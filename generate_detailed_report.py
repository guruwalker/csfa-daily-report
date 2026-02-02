"""
Refactored CSFA Report Generator with Monthly Attendance Tracking
Generates detailed Excel reports with visits, orders, and product details.
Filters out test accounts and tracks monthly attendance using attendance.json
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
from attendance_tracker import AttendanceTracker, update_attendance_for_date, generate_monthly_attendance_summary

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
    customer_fill_color: str = "4BACC6"  # Bright blue
    product_header_color: str = "92D050"  # Light green
    error_color: str = "FF0000"
    perfect_attendance_color: str = "C6EFCE"  # Light green
    absent_color: str = "FFC7CE"  # Light red

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

            # Skip excluded accounts
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

            # Skip excluded accounts
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
        # Merge by ERP code
        df_merge_code = pd.merge(
            df_visits,
            df_orders,
            left_on="erp_code",
            right_on="customer_code",
            how="left",
            suffixes=("_visit", "_order")
        )

        # Merge by customer name
        df_merge_name = pd.merge(
            df_visits,
            df_orders,
            left_on="customer_name",
            right_on="customer_name",
            how="left",
            suffixes=("_visit", "_order")
        )

        # Combine both merges
        df_final = df_merge_code.copy()
        for col in ["order_value", "customer_code", "sales_rep_order", "order_id"]:
            if col in df_merge_name.columns:
                df_final[col] = df_final[col].combine_first(df_merge_name[col])

        # Create final columns
        df_final["customer_name_final"] = df_final.get("customer_name_visit", pd.Series("")).combine_first(
            df_final.get("customer_name_order", pd.Series(""))
        )
        df_final["sales_rep_final"] = df_final.get("sales_rep_visit", pd.Series("")).combine_first(
            df_final.get("sales_rep_order", pd.Series(""))
        )

        return df_final

    @staticmethod
    def get_called_customers(
        df_visits: pd.DataFrame,
        df_orders: pd.DataFrame
    ) -> pd.DataFrame:
        """Find customers who were called but not visited."""
        visited_customers = set(df_visits["customer_name"])
        df_called = df_orders[~df_orders["customer_name"].isin(visited_customers)].copy()
        df_called["customer_called"] = df_called["customer_name"]
        return df_called

    @staticmethod
    def get_sales_reps(
        df_final: pd.DataFrame,
        df_called: pd.DataFrame
    ) -> List[str]:
        """
        Get complete list of all sales representatives.
        Includes both active reps (from data) and inactive reps (from config).
        """
        # Get reps who had activity
        active_reps = set(df_final["sales_rep_final"].dropna()).union(
            df_called["sales_rep"].dropna()
        )

        # Filter out any excluded accounts that might have slipped through
        active_reps = {rep for rep in active_reps if not is_excluded_account(rep)}

        # Get all configured salespeople
        all_configured_reps = set(get_all_salespeople())

        # Combine and sort
        all_reps = active_reps.union(all_configured_reps)

        return sorted(all_reps)


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

    def apply_summary_styling(self, ws: Worksheet) -> None:
        """Apply styling to summary sheet (wrap text, wider columns, with borders)."""
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
            for col_idx, cell in enumerate(ws[1], start=1):
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border
        except IndexError:
            logger.warning("Cannot style header row - worksheet may be empty")
            return

        body_font = Font(
            name=self.config.body_font_name,
            size=self.config.body_font_size
        )
        body_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.font = body_font
                cell.alignment = body_align
                cell.border = thin_border

        column_widths = {
            1: 25,  # SALESPERSON
            2: 20,  # CUSTOMERS VISITED
            3: 30,  # ORDER VALUE FROM VISITS
            4: 20,  # CUSTOMERS CALLED
            5: 30,  # ORDER VALUE FROM CALLS
            6: 20,  # APP USAGE
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

    def apply_attendance_styling(self, ws: Worksheet) -> None:
        """Apply styling to monthly attendance sheet with color coding."""
        if not ws or ws.max_row == 0:
            logger.warning("Empty worksheet, skipping styling")
            return

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Header styling
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

        # Body styling with conditional formatting
        perfect_fill = PatternFill(
            start_color=self.config.perfect_attendance_color,
            end_color=self.config.perfect_attendance_color,
            fill_type="solid"
        )
        absent_fill = PatternFill(
            start_color=self.config.absent_color,
            end_color=self.config.absent_color,
            fill_type="solid"
        )

        body_font = Font(name=self.config.body_font_name, size=self.config.body_font_size)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            attendance_cell = row[1] if len(row) > 1 else None  # Attendance column

            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell.border = thin_border

                # Color code based on attendance status
                if attendance_cell and attendance_cell.value:
                    if "Perfect attendance" in str(attendance_cell.value):
                        cell.fill = perfect_fill
                    elif "No attendance" in str(attendance_cell.value):
                        cell.fill = absent_fill

        # Set column widths
        ws.column_dimensions['A'].width = 30  # Salesperson
        ws.column_dimensions['B'].width = 60  # Attendance

        # Adjust row heights
        for row_idx in range(2, ws.max_row + 1):
            attendance_cell = ws.cell(row=row_idx, column=2)
            if attendance_cell.value:
                height = self.calculate_row_height(str(attendance_cell.value), 60, self.config.body_font_size)
                ws.row_dimensions[row_idx].height = max(25, height)

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
        df_called: pd.DataFrame
    ) -> pd.DataFrame:
        """Generate summary statistics for each sales rep."""
        summary_rows = []

        for rep in reps:
            # Visits data
            rep_visits = df_final[df_final["sales_rep_final"] == rep]
            customers_visited = rep_visits["customer_name_final"].nunique()

            # Calls data
            rep_calls = df_called[df_called["sales_rep"] == rep]
            customers_called = rep_calls["customer_called"].nunique()

            # Determine app usage
            app_usage = "Used App" if (customers_visited > 0 or customers_called > 0) else "Did Not Use App"

            summary_rows.append({
                "SALESPERSON": rep,
                "CUSTOMERS VISITED": customers_visited,
                "CUSTOMERS CALLED": customers_called,
                "APP USAGE": app_usage
            })

        return pd.DataFrame(summary_rows)

    def export_summary_text(self, df_summary: pd.DataFrame, filepath: str) -> None:
        """Export summary as formatted text for email."""
        summary_for_email = df_summary.copy()

        summary_for_email["CUSTOMERS VISITED"] = \
            summary_for_email["CUSTOMERS VISITED"].map(lambda x: f"{int(x):,}")
        summary_for_email["CUSTOMERS CALLED"] = \
            summary_for_email["CUSTOMERS CALLED"].map(lambda x: f"{int(x):,}")

        with open(filepath, "w") as f:
            f.write(summary_for_email.to_string(index=False))

        logger.info(f"✅ Summary text saved: {filepath}")

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
# REP SHEET GENERATOR
# ============================================================================

class RepSheetGenerator:
    """Generates individual sales rep sheets."""

    def __init__(self, access_token: str, config: ReportConfig):
        self.access_token = access_token
        self.config = config

        self.customer_fill = PatternFill(
            start_color=config.customer_fill_color,
            end_color=config.customer_fill_color,
            fill_type="solid"
        )
        self.product_header_fill = PatternFill(
            start_color=config.product_header_color,
            end_color=config.product_header_color,
            fill_type="solid"
        )

    def create_rep_sheet(
        self,
        writer: pd.ExcelWriter,
        rep: str,
        df_final: pd.DataFrame,
        df_called: pd.DataFrame,
        orders_data: List[Dict]
    ) -> None:
        """Create a complete sheet for a sales rep."""
        logger.info(f"  Processing: {rep}")

        rep_visits = df_final[df_final["sales_rep_final"] == rep]
        rep_calls = df_called[df_called["sales_rep"] == rep]
        customers = self._build_customer_dict(rep, rep_visits, rep_calls, orders_data)

        sheet_name = rep.replace(".", "_").replace(" ", "_")[:31]

        writer.book.create_sheet(sheet_name)
        ws = writer.book[sheet_name]

        self._write_rep_data(ws, customers, rep)
        self._adjust_column_widths_and_heights(ws)

    def _build_customer_dict(
        self,
        rep: str,
        rep_visits: pd.DataFrame,
        rep_calls: pd.DataFrame,
        orders_data: List[Dict]
    ) -> Dict[str, Dict]:
        """Build dictionary of customer information."""
        customers = {}

        for _, visit in rep_visits.iterrows():
            cust = visit["customer_name_final"]
            customers[cust] = {
                "visit_type": "Visited",
                "time_spent": visit.get("time_spent", ""),
                "orders": [],
            }

        for _, call in rep_calls.iterrows():
            cust = call["customer_called"]
            customers.setdefault(
                cust,
                {"visit_type": "Called", "time_spent": "", "orders": []},
            )
            if customers[cust]["visit_type"] == "Visited":
                customers[cust]["visit_type"] = "Visited & Called"

        for order in orders_data:
            if order.get("sales_rep") != rep:
                continue

            cust = order.get("customer_name")
            if not cust:
                continue

            customers.setdefault(
                cust,
                {"visit_type": "No Visit", "time_spent": "", "orders": []},
            )
            customers[cust]["orders"].append(order.get("id"))

        return customers

    def _write_rep_data(self, ws: Worksheet, customers: Dict[str, Dict], rep: str) -> None:
        """Write customer data to worksheet with styling."""
        row_idx = 1
        wrap_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for cust_name, info in customers.items():
            visit_type = info["visit_type"]
            if visit_type == "Visited":
                customer_display = f"{cust_name} (visited)"
            elif visit_type == "Called":
                customer_display = f"{cust_name} (called)"
            elif visit_type == "Visited & Called":
                customer_display = f"{cust_name} (visited & called)"
            else:
                customer_display = cust_name

            ws.cell(row=row_idx, column=1, value=customer_display)

            time_spent = info["time_spent"]
            if time_spent and visit_type in ["Visited", "Visited & Called"]:
                time_display = f"Time Spent: {time_spent}"
            else:
                time_display = ""

            ws.cell(row=row_idx, column=2, value=time_display)
            ws.cell(row=row_idx, column=3, value="")
            ws.cell(row=row_idx, column=4, value="")
            ws.cell(row=row_idx, column=5, value="")
            ws.cell(row=row_idx, column=6, value="")
            ws.cell(row=row_idx, column=7, value="")

            for col_idx in range(1, 8):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.fill = self.customer_fill
                cell.font = Font(
                    name=self.config.body_font_name,
                    size=self.config.body_font_size,
                    bold=True
                )
                cell.alignment = wrap_align
                cell.border = thin_border

            row_idx += 1

            all_items = self._fetch_order_items(info["orders"], rep)

            if all_items:
                ws.cell(row=row_idx, column=1, value="Product ID")
                ws.cell(row=row_idx, column=2, value="Product Description")
                ws.cell(row=row_idx, column=3, value="")
                ws.cell(row=row_idx, column=4, value="Sold Qty")
                ws.cell(row=row_idx, column=5, value="Unit Cost")
                ws.cell(row=row_idx, column=6, value="Order Value")
                ws.cell(row=row_idx, column=7, value="")

                for col_idx in range(1, 8):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.fill = self.product_header_fill
                    cell.font = Font(
                        name=self.config.body_font_name,
                        size=self.config.body_font_size,
                        bold=True,
                        color="FFFFFF"
                    )
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = thin_border

                row_idx += 1

                for item in all_items:
                    qty = float(item.get("sold_qty", 0))
                    cost = float(item.get("unit_cost", 0))

                    product_desc = item.get("product_desc", "")
                    product_id = str(item.get("product_id", ""))

                    if product_desc and product_id:
                        prefix = f"{product_id} - "
                        if product_desc.startswith(prefix):
                            product_desc = product_desc[len(prefix):]

                    ws.cell(row=row_idx, column=1, value=product_id)
                    ws.cell(row=row_idx, column=2, value=product_desc)
                    ws.cell(row=row_idx, column=3, value="")
                    ws.cell(row=row_idx, column=4, value=qty)
                    ws.cell(row=row_idx, column=5, value=cost)
                    ws.cell(row=row_idx, column=6, value=f"{qty * cost:,.2f}")
                    ws.cell(row=row_idx, column=7, value="")

                    for col_idx in range(1, 8):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.font = Font(
                            name=self.config.body_font_name,
                            size=self.config.body_font_size
                        )
                        cell.alignment = wrap_align
                        cell.border = thin_border

                    row_idx += 1
            else:
                no_orders_cell = ws.cell(row=row_idx, column=1)
                no_orders_cell.value = "No orders"
                no_orders_cell.font = Font(
                    name=self.config.body_font_name,
                    size=self.config.body_font_size,
                    color=self.config.error_color,
                    bold=True
                )
                no_orders_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                no_orders_cell.border = thin_border

                for col_idx in range(1, 8):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.border = thin_border

                ws.merge_cells(
                    start_row=row_idx,
                    start_column=1,
                    end_row=row_idx,
                    end_column=7
                )

                row_idx += 1

            for col_idx in range(1, 8):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = thin_border

            row_idx += 1

    def _fetch_order_items(self, order_ids: List[int], rep: str) -> List[Dict]:
        """Fetch items for all order IDs."""
        all_items = []

        for order_id in order_ids:
            try:
                details = get_order_details(self.access_token, order_id)
                all_items.extend(details.get("entries", []))
            except Exception as e:
                logger.error(f"Error fetching order {order_id} for {rep}: {e}")

        return all_items

    def _adjust_column_widths_and_heights(self, ws: Worksheet) -> None:
        """Set wider column widths and calculate row heights for rep sheet."""
        column_widths = {
            1: 45,
            2: 40,
            3: 20,
            4: 15,
            5: 15,
            6: 18,
            7: 10,
        }

        for col_idx, width in column_widths.items():
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        styler = ExcelStyler(self.config)
        for row_idx in range(1, ws.max_row + 1):
            max_height = 18

            for col_idx in range(1, 8):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell_value = cell.value

                if cell_value is None or (isinstance(cell_value, str) and not cell_value.strip()):
                    continue

                col_width = column_widths.get(col_idx, 20)

                try:
                    height = styler.calculate_row_height(cell_value, col_width, self.config.body_font_size)
                    max_height = max(max_height, height)
                except Exception as e:
                    logger.warning(f"Could not calculate height for row {row_idx}, col {col_idx}: {e}")
                    max_height = max(max_height, 35)

            ws.row_dimensions[row_idx].height = max_height


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
    Generate detailed CSFA report with visits, orders, and monthly attendance tracking.

    Args:
        visits_data: List of visit records
        orders_data: List of order records
        report_date: Date for which the report is being generated (defaults to today)
        config: Optional configuration object
    """
    if config is None:
        config = ReportConfig.from_env()

    if report_date is None:
        report_date = datetime.now()

    access_token = os.getenv("ACCESS_TOKEN")
    if not access_token:
        raise ValueError("ACCESS_TOKEN not found in environment variables")

    logger.info(f"🚀 Starting report generation: {config.output_file}")
    logger.info(f"📅 Report date: {report_date.strftime('%Y-%m-%d')}")

    # Remove old file
    if os.path.exists(config.output_file):
        os.remove(config.output_file)
        logger.info(f"Removed old file: {config.output_file}")

    # Initialize processors
    processor = DataProcessor()
    styler = ExcelStyler(config)
    summary_gen = SummaryGenerator()
    rep_gen = RepSheetGenerator(access_token, config)

    # Initialize attendance tracker
    attendance_tracker = AttendanceTracker(config.attendance_json_file)

    # Process data
    logger.info("📊 Processing data...")
    df_visits = processor.clean_visits(visits_data)
    df_orders = processor.clean_orders(orders_data)
    df_final = processor.merge_visits_orders(df_visits, df_orders)
    df_called = processor.get_called_customers(df_visits, df_orders)
    reps = processor.get_sales_reps(df_final, df_called)

    logger.info(f"Found {len(reps)} sales representatives")

    # Generate summary
    df_summary = summary_gen.generate_summary(reps, df_final, df_called)

    # Update attendance tracking
    logger.info("📅 Updating attendance tracking...")
    salespeople_present = list(df_summary[df_summary["APP USAGE"] == "Used App"]["SALESPERSON"])
    update_attendance_for_date(attendance_tracker, salespeople_present, reps, report_date)

    # Generate monthly attendance summary
    month_name = report_date.strftime("%B")  # e.g., "February"
    logger.info(f"📋 Generating monthly attendance summary for {month_name}...")
    attendance_summary = generate_monthly_attendance_summary(
        attendance_tracker,
        reps,
        report_date.year,
        report_date.month
    )
    df_attendance = pd.DataFrame(attendance_summary)

    # Create Excel file
    logger.info("📝 Creating Excel file...")
    with pd.ExcelWriter(config.output_file, engine="openpyxl") as writer:
        # 1. Summary sheet
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        ws = writer.sheets["Summary"]
        styler.apply_summary_styling(ws)
        styler.format_money_columns(ws, [2, 4])

        # 2. Monthly Attendance sheet
        df_attendance.to_excel(writer, index=False, sheet_name=f"Monthly Attendance - {month_name}")
        ws_attendance = writer.sheets[f"Monthly Attendance - {month_name}"]
        styler.apply_attendance_styling(ws_attendance)

        # 3. Individual rep sheets (only for reps who used the app)
        for rep in reps:
            rep_data = df_summary[df_summary["SALESPERSON"] == rep]
            if not rep_data.empty and rep_data.iloc[0]["APP USAGE"] == "Used App":
                rep_gen.create_rep_sheet(writer, rep, df_final, df_called, orders_data)
            else:
                logger.info(f"  Skipping {rep} (did not use app)")

    # Export summary files
    summary_gen.export_summary_text(df_summary, config.summary_text_file)
    summary_gen.export_summary_image(df_summary, config.summary_image_file)

    logger.info(f"✅ Report generation complete: {config.output_file}")
    logger.info(f"✅ Attendance tracking updated in: {config.attendance_json_file}")


# ============================================================================
# EXAMPLE USAGE
# ============================================================================

if __name__ == "__main__":
    logger.info("Report generator module with monthly attendance tracking loaded successfully")
