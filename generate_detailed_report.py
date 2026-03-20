"""
Simplified CSFA Report Generator with Monthly Attendance Tracking
Generates Excel reports with visits, orders, and attendance.
Focuses on customer visits only (no "customers called" metric).
Simplified structure: Summary + Attendance sheets only (no individual rep sheets).

Updated: Supports suspended salespeople and new starters.
"""

import pandas as pd
import os
import logging
from typing import List, Dict, Any, Set, Tuple
from dataclasses import dataclass
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# Optional: dataframe_image for summary export
try:
    import dataframe_image as dfi
    HAS_DFI = True
except ImportError:
    HAS_DFI = False
    logging.warning("dataframe_image not installed. Summary image export disabled.")

from api_client import get_order_details
from config import is_excluded_account, get_all_salespeople, normalize_name
from attendance_tracker import (
    AttendanceTracker,
    update_attendance_for_date,
    generate_monthly_attendance_summary,
    generate_monthly_attendance_grid,
)
from date_utils import get_report_date
from holiday_checker import HolidayChecker
from leave_checker import LeaveChecker
from suspension_checker import SuspensionChecker
from new_starter_checker import NewStarterChecker

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ============================================================================
# CONFIGURATION
# ============================================================================

@dataclass
class ReportConfig:
    """Configuration for report generation."""
    output_file: str = "Daily_CSFA_Report.xlsx"
    summary_text_file: str = "summary_for_email.txt"
    summary_image_file: str = "summary_sheet.png"
    attendance_json_file: str = "attendance.json"

    # Styling colors
    header_color: str = "4F81BD"
    customer_fill_color: str = "4BACC6"
    product_header_color: str = "92D050"
    error_color: str = "FF0000"
    perfect_attendance_color: str = "C6EFCE"
    absent_color: str = "FFC7CE"
    suspended_color: str = "D9D9D9"  # Light grey for suspended rows

    # Fonts
    header_font_name: str = "Times New Roman"
    header_font_size: int = 12
    body_font_name: str = "Times New Roman"
    body_font_size: int = 12

    @classmethod
    def from_env(cls):
        """Create config from environment variables."""
        return cls(
            output_file=os.getenv("OUTPUT_FILE", "Daily_CSFA_Report.xlsx"),
            summary_text_file=os.getenv("SUMMARY_TEXT_FILE", "summary_for_email.txt"),
            summary_image_file=os.getenv("SUMMARY_IMAGE_FILE", "summary_sheet.png"),
            attendance_json_file=os.getenv("ATTENDANCE_JSON_FILE", "attendance.json"),
        )

    def _side_file(self, filename: str) -> str:
        """Return a path in the same directory as attendance_json_file."""
        base_dir = os.path.dirname(self.attendance_json_file) or "."
        return os.path.join(base_dir, filename)

    @property
    def suspended_json_file(self) -> str:
        return os.getenv("SUSPENDED_FILE", self._side_file("suspended_salespeople.json"))

    @property
    def new_starters_json_file(self) -> str:
        return os.getenv("NEW_STARTERS_FILE", self._side_file("new_starters.json"))


# ============================================================================
# DATA PROCESSING
# ============================================================================

class DataProcessor:
    """Handles data cleaning and processing."""

    @staticmethod
    def clean_visits(visits_data: List[Dict]) -> pd.DataFrame:
        """Clean and structure visits data, filtering out excluded accounts."""
        visit_rows = []
        for v in visits_data:
            rep_name = v.get("rep_name", "")

            if is_excluded_account(rep_name):
                logger.debug(f"Filtering out excluded visit from: {rep_name}")
                continue

            erp_code = v.get("erp_code") or ""
            visit_rows.append({
                "sales_rep": normalize_name(rep_name),
                "customer_name": (v.get("shop_name") or "").strip(),
                "erp_code": erp_code.strip() if erp_code else "",
                "time_spent": v.get("timespent", "")
            })
        return pd.DataFrame(visit_rows)

    @staticmethod
    def clean_orders(orders_data: List[Dict]) -> pd.DataFrame:
        """Clean and structure orders data, filtering out excluded accounts."""
        order_rows = []
        for o in orders_data:
            sales_rep = o.get("sales_rep", "")

            if is_excluded_account(sales_rep):
                logger.debug(f"Filtering out excluded order from: {sales_rep}")
                continue

            balance_str = o.get("balance", "0").replace(",", "").strip()
            try:
                balance = float(balance_str)
            except (ValueError, TypeError):
                balance = 0.0

            order_rows.append({
                "sales_rep": normalize_name(sales_rep),
                "customer_name": (o.get("customer_name") or "").strip(),
                "customer_code": o.get("customer_code", ""),
                "order_id": o.get("id"),
                "order_value": balance
            })
        return pd.DataFrame(order_rows)

    @staticmethod
    def merge_visits_orders(
        df_visits: pd.DataFrame,
        df_orders: pd.DataFrame
    ) -> pd.DataFrame:
        """Merge visits and orders data with fallback logic."""
        if df_visits.empty and df_orders.empty:
            return pd.DataFrame(columns=[
                "customer_name_final",
                "sales_rep_final",
                "time_spent",
                "order_value"
            ])

        if df_visits.empty:
            df_final = df_orders.copy()
            df_final["customer_name_final"] = df_final["customer_name"]
            df_final["sales_rep_final"] = df_final["sales_rep"]
            df_final["time_spent"] = ""
            return df_final

        if df_orders.empty:
            df_final = df_visits.copy()
            df_final["customer_name_final"] = df_final["customer_name"]
            df_final["sales_rep_final"] = df_final["sales_rep"]
            df_final["order_value"] = 0.0
            return df_final

        df_merge_code = pd.merge(
            df_visits,
            df_orders,
            left_on="erp_code",
            right_on="customer_code",
            how="left",
            suffixes=("_visit", "_order")
        )

        df_merge_name = pd.merge(
            df_visits,
            df_orders,
            left_on="customer_name",
            right_on="customer_name",
            how="left",
            suffixes=("_visit", "_order")
        )

        df_final = df_merge_code.copy()
        for col in ["order_value", "customer_code", "sales_rep_order", "order_id"]:
            if col in df_merge_name.columns:
                df_final[col] = df_final[col].combine_first(df_merge_name[col])

        df_final["customer_name_final"] = df_final.get("customer_name_visit", pd.Series("")).combine_first(
            df_final.get("customer_name_order", pd.Series(""))
        )
        df_final["sales_rep_final"] = df_final.get("sales_rep_visit", pd.Series("")).combine_first(
            df_final.get("sales_rep_order", pd.Series(""))
        )

        return df_final

    @staticmethod
    def get_sales_reps(df_final: pd.DataFrame) -> List[str]:
        """Get complete list of all sales representatives."""
        active_reps = set(df_final["sales_rep_final"].dropna())
        active_reps = {rep for rep in active_reps if not is_excluded_account(rep)}
        all_configured_reps = set(get_all_salespeople())
        all_reps = active_reps.union(all_configured_reps)
        return sorted(all_reps)


# ============================================================================
# DAY VISITS GENERATOR
# ============================================================================

def _generate_day_visits_sheet(df_final: pd.DataFrame, reps: List[str]) -> pd.DataFrame:
    """Generate Day Visits sheet showing all customer interactions."""
    visit_rows = []

    for rep in sorted(reps):
        rep_visits = df_final[df_final["sales_rep_final"] == rep]

        if rep_visits.empty:
            continue

        for _, visit in rep_visits.iterrows():
            customer_name = visit["customer_name_final"]
            time_spent = visit.get("time_spent", "")

            visit_rows.append({
                "Salesperson": rep,
                "Customer Name": customer_name,
                "Time Spent": time_spent if time_spent else "-"
            })

    return pd.DataFrame(visit_rows)


# ============================================================================
# EXCEL STYLING
# ============================================================================

class ExcelStyler:
    """Handles Excel worksheet styling."""

    def __init__(self, config: ReportConfig):
        self.config = config

    def calculate_row_height(self, text: str, column_width: float, font_size: int = 12) -> float:
        """Calculate the required row height for wrapped text."""
        if not text or pd.isna(text):
            return 15

        text = str(text)
        chars_per_line = max(1, int(column_width * 0.85))
        lines = text.split('\n')
        total_lines = 0

        for line in lines:
            if len(line) == 0:
                total_lines += 1
            else:
                line_count = max(1, int(len(line) / chars_per_line) + (1 if len(line) % chars_per_line > 0 else 0))
                total_lines += line_count

        line_height = font_size * 1.3
        calculated_height = (total_lines * line_height) + 8

        return max(18, min(409, calculated_height))

    def apply_summary_styling(self, ws: Worksheet, suspension_checker: SuspensionChecker = None) -> None:
        """Apply styling to summary sheet. Suspended rows get grey fill."""
        if not ws or ws.max_row == 0:
            logger.warning("Empty worksheet, skipping styling")
            return

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        header_font = Font(
            name=self.config.header_font_name,
            size=self.config.header_font_size,
            bold=True,
            color="FFFFFF"
        )
        header_fill = PatternFill(
            start_color=self.config.header_color,
            end_color=self.config.header_color,
            fill_type="solid"
        )
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        try:
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border
        except IndexError:
            logger.warning("Cannot style header row — worksheet may be empty")
            return

        body_font = Font(name=self.config.body_font_name, size=self.config.body_font_size)
        body_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        suspended_fill = PatternFill(
            start_color=self.config.suspended_color,
            end_color=self.config.suspended_color,
            fill_type="solid"
        )

        # Find the SALESPERSON and APP USAGE column indices (1-based)
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        salesperson_col = next((i + 1 for i, h in enumerate(headers) if h == "SALESPERSON"), 1)
        app_usage_col = next((i + 1 for i, h in enumerate(headers) if h == "APP USAGE"), None)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            # Check if this row is a suspended person
            is_suspended_row = False
            if app_usage_col:
                app_usage_cell = row[app_usage_col - 1]
                if app_usage_cell.value == "Suspended":
                    is_suspended_row = True

            for cell in row:
                cell.font = body_font
                cell.alignment = body_align
                cell.border = thin_border
                if is_suspended_row:
                    cell.fill = suspended_fill

        column_widths = {
            1: 25,  # SALESPERSON
            2: 20,  # CUSTOMERS VISITED
            3: 20,  # APP USAGE
        }

        for col_idx, width in column_widths.items():
            if col_idx <= ws.max_column:
                column = get_column_letter(col_idx)
                ws.column_dimensions[column].width = width

        ws.row_dimensions[1].height = 30

        for row_idx in range(2, ws.max_row + 1):
            max_height = 18
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell_value = cell.value
                if cell_value is None or (isinstance(cell_value, str) and not cell_value.strip()):
                    continue
                col_width = column_widths.get(col_idx, 20)
                try:
                    height = self.calculate_row_height(cell_value, col_width, self.config.body_font_size)
                    max_height = max(max_height, height)
                except Exception as e:
                    logger.warning(f"Could not calculate height for row {row_idx}, col {col_idx}: {e}")
                    max_height = max(max_height, 35)
            ws.row_dimensions[row_idx].height = max_height

    def apply_attendance_grid_styling(self, ws: Worksheet) -> None:
        """Apply styling to monthly attendance grid."""
        if not ws or ws.max_row == 0:
            logger.warning("Empty worksheet, skipping styling")
            return

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        header_font = Font(
            name=self.config.header_font_name,
            size=self.config.header_font_size,
            bold=True,
            color="FFFFFF"
        )
        header_fill = PatternFill(
            start_color=self.config.header_color,
            end_color=self.config.header_color,
            fill_type="solid"
        )

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        body_font = Font(name=self.config.body_font_name, size=self.config.body_font_size)
        sus_fill = PatternFill(
            start_color=self.config.suspended_color,
            end_color=self.config.suspended_color,
            fill_type="solid"
        )

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for idx, cell in enumerate(row):
                cell.font = body_font
                cell.border = thin_border

                if idx == 0:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Grey fill for SUS cells
                if cell.value == "SUS":
                    cell.fill = sus_fill

        ws.column_dimensions['A'].width = 25

        for col_idx in range(2, ws.max_column):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 4

        summary_col = get_column_letter(ws.max_column)
        ws.column_dimensions[summary_col].width = 35

        ws.freeze_panes = 'B2'

    def apply_attendance_styling(self, ws: Worksheet) -> None:
        """Apply styling to monthly attendance sheet."""
        if not ws or ws.max_row == 0:
            return

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        header_font = Font(
            name=self.config.header_font_name,
            size=self.config.header_font_size,
            bold=True,
            color="FFFFFF"
        )
        header_fill = PatternFill(
            start_color=self.config.header_color,
            end_color=self.config.header_color,
            fill_type="solid"
        )

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        body_font = Font(name=self.config.body_font_name, size=self.config.body_font_size)
        sus_fill = PatternFill(
            start_color=self.config.suspended_color,
            end_color=self.config.suspended_color,
            fill_type="solid"
        )

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            attendance_value = row[1].value if len(row) > 1 else ""
            is_suspended_row = isinstance(attendance_value, str) and "Suspended" in attendance_value

            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell.border = thin_border
                if is_suspended_row:
                    cell.fill = sus_fill

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 60

        for row_idx in range(2, ws.max_row + 1):
            attendance_cell = ws.cell(row=row_idx, column=2)
            if attendance_cell.value:
                height = self.calculate_row_height(str(attendance_cell.value), 60, self.config.body_font_size)
                ws.row_dimensions[row_idx].height = max(25, height)

    def apply_day_visits_styling(self, ws: Worksheet) -> None:
        """Apply styling to Day Visits sheet."""
        if not ws or ws.max_row == 0:
            return

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        header_font = Font(
            name=self.config.header_font_name,
            size=self.config.header_font_size,
            bold=True,
            color="FFFFFF"
        )
        header_fill = PatternFill(
            start_color=self.config.header_color,
            end_color=self.config.header_color,
            fill_type="solid"
        )

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        body_font = Font(name=self.config.body_font_name, size=self.config.body_font_size)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell.border = thin_border

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 20

    def format_money_columns(self, ws: Worksheet, columns: List[int]) -> None:
        """Format specific columns as money."""
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for col_idx in columns:
                if col_idx < len(row):
                    row[col_idx].number_format = '#,##0.00'


# ============================================================================
# SUMMARY GENERATOR
# ============================================================================

class SummaryGenerator:
    """Generates summary reports."""

    def generate_summary(
        self,
        reps: List[str],
        df_final: pd.DataFrame,
        report_date: datetime,
        leave_checker: LeaveChecker,
        suspension_checker: SuspensionChecker = None,
        new_starter_checker: NewStarterChecker = None,
    ) -> pd.DataFrame:
        """Generate summary statistics for each sales rep (visits only)."""
        summary_rows = []

        for rep in reps:
            # Suspended takes top priority
            if suspension_checker and suspension_checker.is_suspended(rep, report_date):
                summary_rows.append({
                    "SALESPERSON": rep,
                    "CUSTOMERS VISITED": 0,
                    "APP USAGE": "Suspended",
                })
                continue

            # Before new starter's start date — show as not yet active
            if new_starter_checker and new_starter_checker.is_before_start(rep, report_date):
                # Shouldn't normally appear on this date, but guard just in case
                summary_rows.append({
                    "SALESPERSON": rep,
                    "CUSTOMERS VISITED": 0,
                    "APP USAGE": "Not yet started",
                })
                continue

            is_on_leave = leave_checker.is_on_leave(rep, report_date)

            rep_visits = df_final[df_final["sales_rep_final"] == rep]
            customers_visited = rep_visits["customer_name_final"].nunique()

            if is_on_leave:
                app_usage = "On Leave"
            elif customers_visited > 0:
                app_usage = "Used App"
            else:
                app_usage = "Did Not Use App"

            summary_rows.append({
                "SALESPERSON": rep,
                "CUSTOMERS VISITED": customers_visited,
                "APP USAGE": app_usage,
            })

        return pd.DataFrame(summary_rows)

    def export_summary_text(self, df_summary: pd.DataFrame, filepath: str) -> None:
        """Export summary as formatted text for email."""
        summary_for_email = df_summary.copy()
        summary_for_email["CUSTOMERS VISITED"] = \
            summary_for_email["CUSTOMERS VISITED"].map(lambda x: f"{int(x):,}")

        with open(filepath, "w", encoding="utf-8") as f:
            f.write(summary_for_email.to_string(index=False))

        logger.info(f"Summary text saved: {filepath}")

    def export_summary_image(self, df_summary: pd.DataFrame, filepath: str) -> None:
        """Export summary as image (optional)."""
        if not HAS_DFI:
            logger.warning("dataframe_image not available, skipping image export")
            return

        try:
            dfi.export(df_summary, filepath)
            logger.info(f"✅ Summary image saved: {filepath}")
        except Exception as e:
            logger.error(f"Failed to export summary image: {e}")


# ============================================================================
# HOLIDAY REPORT GENERATOR
# ============================================================================

def _generate_holiday_report(report_date: datetime, config: ReportConfig) -> None:
    """Generate a special holiday report with centered message."""
    from openpyxl.styles import Font, Alignment

    if os.path.exists(config.output_file):
        os.remove(config.output_file)

    logger.info("📝 Creating holiday report...")

    holiday_message = f"🎉 PUBLIC HOLIDAY - {report_date.strftime('%A, %B %d, %Y')}"
    no_data_message = "No business activities recorded on this date"

    with pd.ExcelWriter(config.output_file, engine="openpyxl") as writer:
        df_holiday = pd.DataFrame({"": [holiday_message, "", no_data_message]})
        df_holiday.to_excel(writer, index=False, sheet_name="Holiday Notice", header=False)

        ws = writer.sheets["Holiday Notice"]

        ws.merge_cells('A1:E1')
        cell_a1 = ws['A1']
        cell_a1.font = Font(name="Times New Roman", size=18, bold=True, color="FF0000")
        cell_a1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        ws.merge_cells('A3:E3')
        cell_a3 = ws['A3']
        cell_a3.font = Font(name="Times New Roman", size=14, italic=True)
        cell_a3.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[3].height = 30

        ws.column_dimensions['A'].width = 100

    logger.info(f"✅ Holiday report generated: {config.output_file}")


# ============================================================================
# MAIN REPORT GENERATOR
# ============================================================================

def generate_detailed_report(
    visits_data: List[Dict],
    orders_data: List[Dict],
    report_date: datetime = None,
    config: ReportConfig = None
) -> None:
    """
    Generate simplified CSFA report with attendance tracking.

    Sheets:
      1. Day Summary
      2. Attendance Grid - {Month}
      3. Monthly Attendance - {Month}
      4. Day Visits

    Supports:
      - Suspended salespeople (shown as "Suspended" in summary, SUS in grid)
      - New starters (blank cells before their start date)
    """
    if config is None:
        config = ReportConfig.from_env()

    if report_date is None:
        report_date = get_report_date()

    access_token = os.getenv("ACCESS_TOKEN")
    if not access_token:
        raise ValueError("ACCESS_TOKEN not found in environment variables")

    logger.info(f"🚀 Starting report generation: {config.output_file}")
    logger.info(f"📅 Report date: {report_date.strftime('%A, %Y-%m-%d')}")
    logger.info(f"⏰ Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Holiday check
    holiday_checker = HolidayChecker(config.attendance_json_file.replace("attendance.json", "holidays.json"))
    if holiday_checker.is_holiday(report_date):
        logger.info("🎉 Holiday detected — generating special holiday report")
        _generate_holiday_report(report_date, config)
        return

    if os.path.exists(config.output_file):
        os.remove(config.output_file)
        logger.info(f"Removed old file: {config.output_file}")

    # Initialise all checkers
    processor = DataProcessor()
    styler = ExcelStyler(config)
    summary_gen = SummaryGenerator()

    attendance_tracker = AttendanceTracker(config.attendance_json_file)
    leave_checker = LeaveChecker(config.attendance_json_file.replace("attendance.json", "leave_days.json"))
    suspension_checker = SuspensionChecker(config.suspended_json_file)
    new_starter_checker = NewStarterChecker(config.new_starters_json_file)

    # Process visit/order data
    logger.info("📊 Processing data...")
    df_visits = processor.clean_visits(visits_data)
    df_orders = processor.clean_orders(orders_data)

    if df_visits.empty and df_orders.empty:
        logger.warning("⚠️  No visits or orders data — generating report with attendance only")

    df_final = processor.merge_visits_orders(df_visits, df_orders)
    reps = processor.get_sales_reps(df_final)

    logger.info(f"Found {len(reps)} sales representatives")

    # Summary (visits only) — includes suspension + new-starter awareness
    df_summary = summary_gen.generate_summary(
        reps, df_final, report_date, leave_checker,
        suspension_checker=suspension_checker,
        new_starter_checker=new_starter_checker,
    )

    # Attendance update — suspended people and pre-start days are skipped
    logger.info("📅 Updating attendance tracking...")
    salespeople_present = list(df_summary[df_summary["APP USAGE"] == "Used App"]["SALESPERSON"])
    update_attendance_for_date(
        attendance_tracker,
        salespeople_present,
        reps,
        report_date,
        leave_checker=leave_checker,
        suspension_checker=suspension_checker,
        new_starter_checker=new_starter_checker,
    )

    month_name = report_date.strftime("%B")

    # Monthly attendance summary (text, for email)
    logger.info(f"📋 Generating monthly attendance summary for {month_name}...")
    attendance_summary = generate_monthly_attendance_summary(
        attendance_tracker,
        reps,
        report_date.year,
        report_date.month,
        leave_checker=leave_checker,
        suspension_checker=suspension_checker,
        new_starter_checker=new_starter_checker,
    )
    df_attendance_summary = pd.DataFrame(attendance_summary)

    # Daily attendance grid (Excel)
    logger.info("📊 Generating daily attendance grid...")
    df_attendance_grid = generate_monthly_attendance_grid(
        attendance_tracker,
        reps,
        report_date.year,
        report_date.month,
        leave_checker=leave_checker,
        current_date=report_date,
        suspension_checker=suspension_checker,
        new_starter_checker=new_starter_checker,
    )

    # Day visits sheet
    logger.info("📋 Generating Day Visits sheet...")
    df_day_visits = _generate_day_visits_sheet(df_final, reps)

    # Write Excel
    logger.info("📝 Creating Excel file...")
    with pd.ExcelWriter(config.output_file, engine="openpyxl") as writer:
        # 1. Day Summary
        df_summary.to_excel(writer, index=False, sheet_name="Day Summary")
        ws = writer.sheets["Day Summary"]
        styler.apply_summary_styling(ws, suspension_checker=suspension_checker)

        # 2. Attendance Grid
        df_attendance_grid.to_excel(writer, index=False, sheet_name=f"Attendance Grid - {month_name}")
        ws_grid = writer.sheets[f"Attendance Grid - {month_name}"]
        styler.apply_attendance_grid_styling(ws_grid)

        # 3. Monthly Attendance Summary
        df_attendance_summary.to_excel(writer, index=False, sheet_name=f"Monthly Attendance - {month_name}")
        ws_att = writer.sheets[f"Monthly Attendance - {month_name}"]
        styler.apply_attendance_styling(ws_att)

        # 4. Day Visits
        df_day_visits.to_excel(writer, index=False, sheet_name="Day Visits")
        ws_visits = writer.sheets["Day Visits"]
        styler.apply_day_visits_styling(ws_visits)

    summary_gen.export_summary_text(df_summary, config.summary_text_file)
    summary_gen.export_summary_image(df_summary, config.summary_image_file)

    logger.info(f"✅ Report generation complete: {config.output_file}")
    logger.info("✅ Sheets: Day Summary | Attendance Grid | Monthly Attendance | Day Visits")
    logger.info(f"✅ Attendance tracking updated: {config.attendance_json_file}")


# ============================================================================
# EXAMPLE USAGE
# ============================================================================

if __name__ == "__main__":
    logger.info("Simplified report generator module loaded successfully")
