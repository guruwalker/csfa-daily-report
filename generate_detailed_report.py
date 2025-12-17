"""
Enhanced CSFA Report Generator
Generates detailed Excel reports with professional styling and comprehensive data.
"""

import os
import logging
from typing import List, Dict, Optional
from dataclasses import dataclass

import pandas as pd
import dataframe_image as dfi
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from api_client import get_order_details

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
    summary_sheet: str = "Summary"
    access_token: str = ""

    @classmethod
    def from_env(cls) -> 'ReportConfig':
        """Create configuration from environment variables."""
        return cls(
            output_file=os.getenv("OUTPUT_FILE", "Daily_CSFA_Report.xlsx"),
            summary_text_file=os.getenv("SUMMARY_TEXT_FILE", "summary_for_email.txt"),
            summary_image_file=os.getenv("SUMMARY_IMAGE_FILE", "summary_sheet.png"),
            summary_sheet=os.getenv("SUMMARY_SHEET", "Summary"),
            access_token=os.getenv("ACCESS_TOKEN", "")
        )


# ============================================================================
# STYLING CONSTANTS
# ============================================================================

class ExcelStyles:
    """Excel styling constants."""

    # Fonts
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    CUSTOMER_FONT = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    BODY_FONT = Font(name="Calibri", size=11)
    ERROR_FONT = Font(name="Calibri", size=11, color="C00000", bold=True)
    PRODUCT_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    # Fills
    HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    CUSTOMER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    PRODUCT_HEADER_FILL = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    NO_ORDER_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Alignments
    CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")

    # Borders
    THIN_BORDER = Border(
        left=Side(style='thin', color='B4B4B4'),
        right=Side(style='thin', color='B4B4B4'),
        top=Side(style='thin', color='B4B4B4'),
        bottom=Side(style='thin', color='B4B4B4')
    )

    THICK_BORDER = Border(
        left=Side(style='medium', color='4472C4'),
        right=Side(style='medium', color='4472C4'),
        top=Side(style='medium', color='4472C4'),
        bottom=Side(style='medium', color='4472C4')
    )

    # Column definitions for rep sheets
    REP_COLUMNS = [
        {"name": "Product/Customer", "width": 40},
        {"name": "Visit Type", "width": 15},
        {"name": "Time Spent", "width": 12},
        {"name": "Qty", "width": 10},
        {"name": "Unit Cost (MZN)", "width": 15},
        {"name": "Order Value (MZN)", "width": 18},
    ]


# ============================================================================
# DATA PROCESSING
# ============================================================================

class DataProcessor:
    """Process visits and orders data."""

    @staticmethod
    def clean_visits(visits_data: List[Dict]) -> pd.DataFrame:
        """Clean and structure visits data."""
        visit_rows = []
        for v in visits_data:
            erp_code = v.get("erp_code") or ""
            visit_rows.append({
                "sales_rep": v.get("rep_name", ""),
                "customer_name": (v.get("shop_name") or "").strip(),
                "erp_code": erp_code.strip() if erp_code else "",
                "time_spent": v.get("timespent", "")
            })
        return pd.DataFrame(visit_rows)

    @staticmethod
    def clean_orders(orders_data: List[Dict]) -> pd.DataFrame:
        """Clean and structure orders data."""
        order_rows = []
        for o in orders_data:
            balance_str = o.get("balance", "0").replace(",", "").strip()
            try:
                balance = float(balance_str)
            except:
                balance = 0.0
            order_rows.append({
                "sales_rep": o.get("sales_rep", ""),
                "customer_name": (o.get("customer_name") or "").strip(),
                "customer_code": o.get("customer_code", ""),
                "order_id": o.get("id"),
                "order_value": balance
            })
        return pd.DataFrame(order_rows)

    @staticmethod
    def merge_data(df_visits: pd.DataFrame, df_orders: pd.DataFrame) -> pd.DataFrame:
        """Merge visits and orders data."""
        # Merge by customer code
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

        # Combine results
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
    def find_called_customers(df_visits: pd.DataFrame, df_orders: pd.DataFrame) -> pd.DataFrame:
        """Find customers who were called but not visited."""
        visited_customers = set(df_visits["customer_name"])
        df_called = df_orders[~df_orders["customer_name"].isin(visited_customers)].copy()
        df_called["customer_called"] = df_called["customer_name"]
        return df_called


# ============================================================================
# EXCEL GENERATOR
# ============================================================================

class ExcelGenerator:
    """Generate Excel reports with professional styling."""

    def __init__(self, config: ReportConfig):
        self.config = config
        self.styles = ExcelStyles()

    def apply_summary_styling(self, ws):
        """Apply styling to summary sheet."""
        # Style header row
        for col_idx, cell in enumerate(ws[1], start=1):
            cell.font = self.styles.HEADER_FONT
            cell.fill = self.styles.HEADER_FILL
            cell.alignment = self.styles.CENTER_ALIGN
            cell.border = self.styles.THIN_BORDER

        # Style body rows
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
            for col_idx, cell in enumerate(row, start=1):
                cell.font = self.styles.BODY_FONT
                cell.border = self.styles.THIN_BORDER

                # Center align text columns, right align numbers
                if col_idx == 1:  # Salesperson column
                    cell.alignment = self.styles.LEFT_ALIGN
                elif col_idx in [2, 4]:  # Customer count columns
                    cell.alignment = self.styles.CENTER_ALIGN
                else:  # Value columns
                    cell.alignment = self.styles.RIGHT_ALIGN

        # Auto-adjust column widths
        column_widths = {
            1: 25,  # Salesperson
            2: 18,  # Customers Visited
            3: 22,  # Order Value from Visits
            4: 18,  # Customers Called
            5: 22,  # Order Value from Calls
        }

        for col_idx, width in column_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Freeze header row
        ws.freeze_panes = ws['A2']

    def create_summary_sheet(self, writer, df_final: pd.DataFrame, df_called: pd.DataFrame, reps: List[str]):
        """Create and style summary sheet."""
        summary_rows = []

        for rep in reps:
            rep_visits = df_final[df_final["sales_rep_final"] == rep]
            customers_visited = rep_visits["customer_name_final"].nunique()
            order_value_visits = rep_visits["order_value"].sum()

            rep_calls = df_called[df_called["sales_rep"] == rep]
            customers_called = rep_calls["customer_called"].nunique()
            order_value_calls = rep_calls["order_value"].sum()

            summary_rows.append({
                "SALESPERSON": rep,
                "CUSTOMERS VISITED": customers_visited,
                "ORDER VALUE FROM VISITS": order_value_visits,
                "CUSTOMERS CALLED": customers_called,
                "ORDER VALUE FROM CALLS": order_value_calls
            })

        df_summary = pd.DataFrame(summary_rows)

        # Write to Excel
        df_summary.to_excel(writer, index=False, sheet_name=self.config.summary_sheet)
        ws = writer.sheets[self.config.summary_sheet]
        self.apply_summary_styling(ws)

        # Format numeric columns as money
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            row[2].number_format = '#,##0.00'  # Order Value from Visits
            row[4].number_format = '#,##0.00'  # Order Value from Calls

        # Save summary files
        self._save_summary_files(df_summary)

        return df_summary

    def _save_summary_files(self, df_summary: pd.DataFrame):
        """Save summary as text and image files."""
        # Text file for email
        summary_for_email = df_summary.copy()
        summary_for_email["ORDER VALUE FROM VISITS"] = summary_for_email["ORDER VALUE FROM VISITS"].map("{:,.2f}".format)
        summary_for_email["ORDER VALUE FROM CALLS"] = summary_for_email["ORDER VALUE FROM CALLS"].map("{:,.2f}".format)

        with open(self.config.summary_text_file, "w") as f:
            f.write(summary_for_email.to_string(index=False))

        logger.info(f"✅ Summary text saved: {self.config.summary_text_file}")

        # Image file
        try:
            dfi.export(df_summary, self.config.summary_image_file)
            logger.info(f"✅ Summary image saved: {self.config.summary_image_file}")
        except Exception as e:
            logger.warning(f"⚠️ Could not save summary image: {e}")

    def create_rep_sheet(self, writer, rep: str, df_final: pd.DataFrame,
                        df_called: pd.DataFrame, orders_data: List[Dict]):
        """Create individual sales rep sheet with professional formatting."""
        rep_visits = df_final[df_final["sales_rep_final"] == rep]
        rep_calls = df_called[df_called["sales_rep"] == rep]

        customers = {}

        # Process visits
        for _, visit in rep_visits.iterrows():
            cust = visit["customer_name_final"]
            customers[cust] = {
                "visit_type": "Visited",
                "time_spent": visit.get("time_spent", ""),
                "orders": [],
            }

        # Process calls
        for _, call in rep_calls.iterrows():
            cust = call["customer_called"]
            customers.setdefault(
                cust,
                {"visit_type": "Called", "time_spent": "", "orders": []},
            )
            if customers[cust]["visit_type"] == "Visited":
                customers[cust]["visit_type"] = "Visited & Called"

        # Process orders
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

        # Create sheet
        sheet_name = rep.replace(".", "_").replace(" ", "_")[:31]
        ws = writer.book.create_sheet(title=sheet_name)

        # Write data with proper formatting
        self._write_rep_data(ws, customers, rep)

    def _write_rep_data(self, ws, customers: Dict, rep_name: str):
        """Write customer data to worksheet with professional styling."""
        current_row = 1

        # Write sheet title
        ws.merge_cells(f'A{current_row}:F{current_row}')
        title_cell = ws.cell(row=current_row, column=1)
        title_cell.value = f"Sales Activity Report - {rep_name}"
        title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
        title_cell.alignment = self.styles.CENTER_ALIGN
        title_cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        current_row += 2

        # Set column widths
        for col_idx, col_info in enumerate(self.styles.REP_COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = col_info["width"]

        # Process each customer
        # for cust_name, info in sorted(customers.items()):
        #     # Customer header row
        #     ws.merge_cells(f'A{current_row}:C{current_row}')
        #     customer_cell = ws.cell(row=current_row, column=1)
        #     customer_cell.value = cust_name
        #     customer_cell.font = self.styles.CUSTOMER_FONT
        #     customer_cell.fill = self.styles.CUSTOMER_FILL
        #     customer_cell.alignment = self.styles.LEFT_ALIGN
        #     customer_cell.border = self.styles.THIN_BORDER
        for cust_name, info in sorted(customers.items()):
            # Customer header row (merge only A–B)
            ws.merge_cells(start_row=current_row, start_column=1,
                        end_row=current_row, end_column=2)

            customer_cell = ws.cell(row=current_row, column=1)
            customer_cell.value = cust_name
            customer_cell.font = self.styles.CUSTOMER_FONT
            customer_cell.fill = self.styles.CUSTOMER_FILL
            customer_cell.alignment = self.styles.LEFT_ALIGN
            customer_cell.border = self.styles.THIN_BORDER

            # Apply fill & border to the merged partner cell (B)
            partner_cell = ws.cell(row=current_row, column=2)
            partner_cell.fill = self.styles.CUSTOMER_FILL
            partner_cell.border = self.styles.THIN_BORDER


            # Visit type
            visit_cell = ws.cell(row=current_row, column=2)
            visit_cell.value = info["visit_type"]
            visit_cell.font = Font(name="Calibri", size=10, italic=True)
            visit_cell.fill = self.styles.CUSTOMER_FILL
            visit_cell.alignment = self.styles.CENTER_ALIGN
            visit_cell.border = self.styles.THIN_BORDER

            # Time spent
            time_cell = ws.cell(row=current_row, column=3)
            time_cell.value = info["time_spent"]
            time_cell.font = Font(name="Calibri", size=10)
            time_cell.fill = self.styles.CUSTOMER_FILL
            time_cell.alignment = self.styles.CENTER_ALIGN
            time_cell.border = self.styles.THIN_BORDER

            # Empty cells for alignment
            for col in range(4, 7):
                cell = ws.cell(row=current_row, column=col)
                cell.fill = self.styles.CUSTOMER_FILL
                cell.border = self.styles.THIN_BORDER

            current_row += 1

            # Get order items
            all_items = []
            for order_id in info["orders"]:
                try:
                    details = get_order_details(self.config.access_token, order_id)
                    all_items.extend(details.get("entries", []))
                except Exception as e:
                    logger.error(f"Error fetching order {order_id}: {e}")

            if all_items:
                # Product header row
                for col_idx, col_info in enumerate(self.styles.REP_COLUMNS, start=1):
                    header_cell = ws.cell(row=current_row, column=col_idx)
                    header_cell.value = col_info["name"]
                    header_cell.font = self.styles.PRODUCT_HEADER_FONT
                    header_cell.fill = self.styles.PRODUCT_HEADER_FILL
                    header_cell.alignment = self.styles.CENTER_ALIGN
                    header_cell.border = self.styles.THIN_BORDER
                current_row += 1

                # Product rows
                for item in all_items:
                    qty = float(item.get("sold_qty", 0))
                    cost = float(item.get("unit_cost", 0))
                    order_value = qty * cost

                    # Product ID
                    prod_cell = ws.cell(row=current_row, column=1)
                    prod_cell.value = item.get("product_id", "")
                    prod_cell.font = self.styles.BODY_FONT
                    prod_cell.alignment = self.styles.LEFT_ALIGN
                    prod_cell.border = self.styles.THIN_BORDER

                    # Empty cells
                    for col in [2, 3]:
                        cell = ws.cell(row=current_row, column=col)
                        cell.value = ""
                        cell.border = self.styles.THIN_BORDER

                    # Quantity
                    qty_cell = ws.cell(row=current_row, column=4)
                    qty_cell.value = qty
                    qty_cell.font = self.styles.BODY_FONT
                    qty_cell.alignment = self.styles.CENTER_ALIGN
                    qty_cell.border = self.styles.THIN_BORDER
                    qty_cell.number_format = '#,##0.00'

                    # Unit Cost
                    cost_cell = ws.cell(row=current_row, column=5)
                    cost_cell.value = cost
                    cost_cell.font = self.styles.BODY_FONT
                    cost_cell.alignment = self.styles.RIGHT_ALIGN
                    cost_cell.border = self.styles.THIN_BORDER
                    cost_cell.number_format = '#,##0.00'

                    # Order Value
                    value_cell = ws.cell(row=current_row, column=6)
                    value_cell.value = order_value
                    value_cell.font = self.styles.BODY_FONT
                    value_cell.alignment = self.styles.RIGHT_ALIGN
                    value_cell.border = self.styles.THIN_BORDER
                    value_cell.number_format = '#,##0.00'

                    current_row += 1
            else:
                # No orders row
                ws.merge_cells(f'A{current_row}:F{current_row}')
                no_order_cell = ws.cell(row=current_row, column=1)
                no_order_cell.value = "No orders for this customer"
                no_order_cell.font = self.styles.ERROR_FONT
                no_order_cell.fill = self.styles.NO_ORDER_FILL
                no_order_cell.alignment = self.styles.CENTER_ALIGN
                no_order_cell.border = self.styles.THIN_BORDER
                current_row += 1

            # Add spacing between customers
            current_row += 1

        # Freeze panes at row 3 (after title and blank row)
        ws.freeze_panes = ws['A3']


# ============================================================================
# MAIN REPORT GENERATOR
# ============================================================================

def generate_detailed_report(visits_data: List[Dict], orders_data: List[Dict],
                            config: Optional[ReportConfig] = None) -> pd.DataFrame:
    """
    Generate detailed CSFA report.

    Args:
        visits_data: List of visit dictionaries
        orders_data: List of order dictionaries
        config: Report configuration (optional)

    Returns:
        DataFrame containing summary data
    """
    if config is None:
        config = ReportConfig.from_env()

    logger.info(f"📊 Generating report: {config.output_file}")

    # Remove old file if exists
    if os.path.exists(config.output_file):
        os.remove(config.output_file)
        logger.info(f"🗑️  Removed old file: {config.output_file}")

    # Process data
    logger.info("🔄 Processing data...")
    processor = DataProcessor()
    df_visits = processor.clean_visits(visits_data)
    df_orders = processor.clean_orders(orders_data)
    df_final = processor.merge_data(df_visits, df_orders)
    df_called = processor.find_called_customers(df_visits, df_orders)

    # Get all sales reps
    reps = sorted(set(df_final["sales_rep_final"].dropna()).union(df_called["sales_rep"].dropna()))
    logger.info(f"👥 Found {len(reps)} sales representatives")

    # Generate Excel
    logger.info("📝 Writing Excel file...")
    generator = ExcelGenerator(config)

    with pd.ExcelWriter(config.output_file, engine="openpyxl") as writer:
        # Create summary sheet
        df_summary = generator.create_summary_sheet(writer, df_final, df_called, reps)

        # Create individual rep sheets
        for idx, rep in enumerate(reps, 1):
            logger.info(f"  [{idx}/{len(reps)}] Creating sheet for: {rep}")
            generator.create_rep_sheet(writer, rep, df_final, df_called, orders_data)

    logger.info(f"✅ Report generated: {config.output_file}")

    return df_summary
