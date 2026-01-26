"""
Enhanced Email Module for CSFA Report
Sends daily reports with professional formatting and error handling.
"""

import os
import smtplib
import logging
from email.message import EmailMessage
from datetime import datetime
from typing import List, Optional
from pathlib import Path
import mimetypes
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)


# ============================================================================
# EMAIL CONFIGURATION
# ============================================================================

class EmailConfig:
    """Email configuration from environment variables."""

    # SMTP Settings
    SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.robbialac.co.mz")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SENDER_EMAIL = os.getenv("SENDER_EMAIL", "innocent.maina@robbialac.co.mz")
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

    # Recipients (comma-separated in .env)
    TO_RECIPIENTS = os.getenv("EMAIL_TO", "innocent.maina@crownpaints.co.ke").split(",")
    CC_RECIPIENTS = os.getenv("EMAIL_CC", "daniel.ndirangu@robbialac.co.mz,isaac.mokua@robbialac.co.mz").split(",")
    BCC_RECIPIENTS = os.getenv("EMAIL_BCC", "").split(",") if os.getenv("EMAIL_BCC") else []

    # Email content
    EMAIL_SUBJECT_TEMPLATE = os.getenv("EMAIL_SUBJECT", "Tintas Berger CSFA Report - {date}")
    SENDER_NAME = os.getenv("SENDER_NAME", "Innocent Maina")
    RECIPIENT_NAME = os.getenv("RECIPIENT_NAME", "Mr. Hussein")

    # Files
    EXCEL_FILE = os.getenv("OUTPUT_FILE", "Daily_CSFA_Report.xlsx")
    SUMMARY_SHEET = os.getenv("SUMMARY_SHEET", "Summary")

    # SMTP timeout
    SMTP_TIMEOUT = int(os.getenv("SMTP_TIMEOUT", "30"))

    @classmethod
    def validate(cls) -> None:
        """Validate required configuration."""
        if not cls.EMAIL_PASSWORD:
            raise ValueError("EMAIL_PASSWORD not set in environment variables")
        if not cls.SENDER_EMAIL:
            raise ValueError("SENDER_EMAIL not set in environment variables")
        if not cls.TO_RECIPIENTS or not cls.TO_RECIPIENTS[0]:
            raise ValueError("EMAIL_TO not set in environment variables")

    @classmethod
    def clean_recipients(cls, recipients: List[str]) -> List[str]:
        """Clean and filter recipient list."""
        return [r.strip() for r in recipients if r.strip()]


# ============================================================================
# HTML TABLE GENERATOR
# ============================================================================

class HTMLTableGenerator:
    """Generate HTML tables from DataFrames."""

    # Styling constants
    HEADER_STYLE = (
        "background-color: #4F81BD; "
        "color: #FFFFFF; "
        "font-weight: bold; "
        "padding: 12px; "
        "text-align: left; "
        "border: 1px solid #2F5F8D; "
        "font-family: Arial, sans-serif;"
    )

    CELL_STYLE = (
        "padding: 10px; "
        "border: 1px solid #ddd; "
        "text-align: left; "
        "font-family: Arial, sans-serif; "
        "color: #333333;"
    )

    TABLE_STYLE = (
        "border-collapse: collapse; "
        "width: 100%; "
        "margin: 20px 0; "
        "box-shadow: 0 2px 4px rgba(0,0,0,0.1);"
    )

    ALT_ROW_STYLE = "background-color: #f9f9f9;"

    @classmethod
    def generate(cls, df: pd.DataFrame, format_money: bool = True) -> str:
        """
        Generate HTML table from DataFrame.

        Args:
            df: DataFrame to convert
            format_money: Whether to format numeric columns as money

        Returns:
            HTML table string
        """
        if df.empty:
            return "<p><em>No data available</em></p>"

        # Format display for specific columns
        df_formatted = df.copy()

        # Format customer count columns as integers (no decimals)
        customer_columns = ['CUSTOMERS VISITED', 'CUSTOMERS CALLED']
        for col in customer_columns:
            if col in df_formatted.columns:
                df_formatted[col] = df_formatted[col].apply(
                    lambda x: f"{int(x):,}" if pd.notnull(x) else ""
                )

        # Format money columns
        if format_money:
            money_columns = ['ORDER VALUE FROM VISITS', 'ORDER VALUE FROM CALLS']
            for col in money_columns:
                if col in df_formatted.columns:
                    df_formatted[col] = df_formatted[col].apply(
                        lambda x: f"{x:,.2f}" if pd.notnull(x) else ""
                    )

        # Build HTML table
        html = f'<table style="{cls.TABLE_STYLE}">'

        # Header row
        html += '<thead><tr>'
        for col in df_formatted.columns:
            html += f'<th style="{cls.HEADER_STYLE}">{col}</th>'
        html += '</tr></thead>'

        # Data rows with alternating colors
        html += '<tbody>'
        for idx, row in df_formatted.iterrows():
            row_style = cls.ALT_ROW_STYLE if idx % 2 == 1 else ""
            html += f'<tr style="{row_style}">'
            for col in df_formatted.columns:
                value = row[col]
                # Right-align numbers
                is_numeric = col in customer_columns or col in money_columns
                align = "right" if is_numeric else "left"
                cell_style = cls.CELL_STYLE + f" text-align: {align};"
                html += f'<td style="{cell_style}">{value}</td>'
            html += '</tr>'
        html += '</tbody>'
        html += '</table>'

        return html


# ============================================================================
# EMAIL BUILDER
# ============================================================================

class EmailBuilder:
    """Build email messages with attachments."""

    def __init__(self, config: EmailConfig):
        self.config = config

    def build_message(
        self,
        summary_html: str,
        date_str: str,
        attachments: Optional[List[str]] = None
    ) -> EmailMessage:
        """
        Build complete email message.

        Args:
            summary_html: HTML table with summary data
            date_str: Date string for subject
            attachments: List of file paths to attach

        Returns:
            EmailMessage object ready to send
        """
        msg = EmailMessage()

        # Set headers
        msg["From"] = self.config.SENDER_EMAIL
        msg["To"] = ", ".join(self.config.clean_recipients(self.config.TO_RECIPIENTS))

        cc_recipients = self.config.clean_recipients(self.config.CC_RECIPIENTS)
        if cc_recipients:
            msg["Cc"] = ", ".join(cc_recipients)

        bcc_recipients = self.config.clean_recipients(self.config.BCC_RECIPIENTS)
        if bcc_recipients:
            msg["Bcc"] = ", ".join(bcc_recipients)

        # Set subject
        subject = self.config.EMAIL_SUBJECT_TEMPLATE.format(date=date_str)
        msg["Subject"] = subject

        # Build HTML body
        body = self._build_html_body(summary_html, date_str)
        msg.add_alternative(body, subtype="html")

        # Add attachments
        if attachments:
            for filepath in attachments:
                self._attach_file(msg, filepath)

        return msg

    def _build_html_body(self, summary_html: str, date_str: str) -> str:
        """Build HTML email body."""
        # Parse date to get day name
        from datetime import datetime
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            day_name = date_obj.strftime("%A")
            formatted_date = f"{day_name}, {date_str}"
        except:
            formatted_date = date_str

        # Get current time for automation message
        current_time = datetime.now().strftime("%I:%M %p")

        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    color: #333;
                }}
                .header {{
                    color: #4F81BD;
                    margin-bottom: 20px;
                }}
                .footer {{
                    margin-top: 30px;
                    color: #666;
                    font-size: 0.9em;
                }}
            </style>
        </head>
        <body>
            <p>Greetings {self.config.RECIPIENT_NAME},</p>

            <p>Please find attached the daily Tintas Berger CSFA report for <strong>{formatted_date}</strong>.</p>

            {summary_html}

            <p>The complete detailed report is attached as an Excel file.</p>

            <div class="footer">
                <p>Kind regards,<br>
                <strong>{self.config.SENDER_NAME}</strong></p>
            </div>
        </body>
        </html>
        """

    def _attach_file(self, msg: EmailMessage, filepath: str) -> None:
        """Attach a file to the email message."""
        if not os.path.exists(filepath):
            logger.warning(f"Attachment not found: {filepath}")
            return

        # Guess MIME type
        ctype, encoding = mimetypes.guess_type(filepath)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"

        maintype, subtype = ctype.split("/", 1)

        # Read and attach file
        try:
            with open(filepath, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=os.path.basename(filepath)
                )
            logger.info(f"✅ Attached: {os.path.basename(filepath)}")
        except Exception as e:
            logger.error(f"❌ Failed to attach {filepath}: {e}")


# ============================================================================
# EMAIL SENDER
# ============================================================================

class EmailSender:
    """Handle SMTP connection and email sending."""

    def __init__(self, config: EmailConfig):
        self.config = config

    def send(self, msg: EmailMessage) -> bool:
        """
        Send email message via SMTP.

        Args:
            msg: EmailMessage to send

        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"📧 Connecting to {self.config.SMTP_SERVER}:{self.config.SMTP_PORT}...")

            with smtplib.SMTP(
                self.config.SMTP_SERVER,
                self.config.SMTP_PORT,
                timeout=self.config.SMTP_TIMEOUT
            ) as server:
                # Enable debug output in development
                if os.getenv("DEBUG", "").lower() == "true":
                    server.set_debuglevel(1)

                # Secure connection
                server.starttls()
                logger.info("🔒 TLS enabled")

                # Login
                server.login(self.config.SENDER_EMAIL, self.config.EMAIL_PASSWORD)
                logger.info(f"✅ Logged in as {self.config.SENDER_EMAIL}")

                # Send
                server.send_message(msg)
                logger.info("✅ Email sent successfully!")

                return True

        except smtplib.SMTPAuthenticationError as e:
            logger.error(f"❌ SMTP Authentication failed: {e}")
            logger.error("Check your EMAIL_PASSWORD in .env file")
            return False

        except smtplib.SMTPException as e:
            logger.error(f"❌ SMTP Error: {e}")
            return False

        except ConnectionError as e:
            logger.error(f"❌ Connection Error: {e}")
            logger.error(f"Cannot connect to {self.config.SMTP_SERVER}:{self.config.SMTP_PORT}")
            return False

        except Exception as e:
            logger.error(f"❌ Unexpected error sending email: {e}")
            return False


# ============================================================================
# CUSTOMER DETAILS TABLE BUILDER
# ============================================================================

def _build_customer_details_table(excel_file: str, df_summary: pd.DataFrame) -> pd.DataFrame:
    """
    Build a customer details table by reading all rep sheets.
    Format: Salesperson, Customer Name, Interaction, Time Spent, Order Value
    """
    from openpyxl import load_workbook

    customer_rows = []

    try:
        # Load the Excel file
        wb = load_workbook(excel_file, data_only=True)

        # Get all salesperson names from summary
        salespersons = df_summary["SALESPERSON"].tolist()

        for rep in salespersons:
            # Clean sheet name (same logic as in report generation)
            sheet_name = rep.replace(".", "_").replace(" ", "_")[:31]

            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]

            # Parse the sheet to extract customer information
            current_customer = None
            current_interaction = None
            current_time_spent = None
            customer_order_value = 0.0

            for row_idx in range(1, ws.max_row + 1):
                # Read first two columns
                col_a = ws.cell(row=row_idx, column=1).value
                col_b = ws.cell(row=row_idx, column=2).value

                if not col_a:
                    continue

                col_a_str = str(col_a).strip()
                col_b_str = str(col_b).strip() if col_b else ""

                # Check if this is a customer header row
                # Customer rows have format: "Customer Name (visited)" or "Customer Name (called)"
                if "(" in col_a_str and col_a_str.endswith(")"):
                    # Save previous customer if exists
                    if current_customer:
                        customer_rows.append({
                            "Salesperson": rep,
                            "Customer Name": current_customer,
                            "Interaction": current_interaction,
                            "Time Spent": current_time_spent or "-",
                            "Order Value": customer_order_value
                        })

                    # Parse new customer
                    # Extract customer name and interaction type
                    if "(visited & called)" in col_a_str.lower():
                        current_customer = col_a_str.rsplit("(", 1)[0].strip()
                        current_interaction = "Visited & Called"
                    elif "(visited)" in col_a_str.lower():
                        current_customer = col_a_str.rsplit("(", 1)[0].strip()
                        current_interaction = "Visited"
                    elif "(called)" in col_a_str.lower():
                        current_customer = col_a_str.rsplit("(", 1)[0].strip()
                        current_interaction = "Called"
                    else:
                        current_customer = col_a_str
                        current_interaction = "Unknown"

                    # Extract time spent from column B
                    if col_b_str.startswith("Time Spent:"):
                        current_time_spent = col_b_str.replace("Time Spent:", "").strip()
                    else:
                        current_time_spent = None

                    # Reset order value for new customer
                    customer_order_value = 0.0

                # Check if this is an order value row (has numeric value in column F)
                elif col_a_str and col_a_str != "Product ID" and col_a_str != "No orders":
                    # Try to get order value from column F (column 6)
                    order_val_cell = ws.cell(row=row_idx, column=6).value
                    if order_val_cell:
                        try:
                            # Remove commas and convert to float
                            order_val_str = str(order_val_cell).replace(",", "")
                            order_val = float(order_val_str)
                            customer_order_value += order_val
                        except (ValueError, AttributeError):
                            pass

            # Don't forget the last customer
            if current_customer:
                customer_rows.append({
                    "Salesperson": rep,
                    "Customer Name": current_customer,
                    "Interaction": current_interaction,
                    "Time Spent": current_time_spent or "-",
                    "Order Value": customer_order_value
                })

        wb.close()

    except Exception as e:
        logger.error(f"Error building customer details table: {e}", exc_info=True)

    # Create DataFrame
    df = pd.DataFrame(customer_rows)

    # Sort by Salesperson, then Customer Name
    if not df.empty:
        df = df.sort_values(by=["Salesperson", "Customer Name"])

    return df


def _generate_customer_details_table(df: pd.DataFrame) -> str:
    """Generate HTML table for customer details."""
    if df.empty:
        return "<p><em>No customer interaction data available</em></p>"

    # Get currency from environment
    currency = os.getenv("REPORT_CURRENCY", "MZN")

    # Define styles
    header_style = (
        "background-color: #4F81BD; "
        "color: #FFFFFF; "
        "font-weight: bold; "
        "padding: 12px; "
        "text-align: left; "
        "border: 1px solid #2F5F8D; "
        "font-family: Arial, sans-serif;"
    )

    cell_style = (
        "padding: 10px; "
        "border: 1px solid #ddd; "
        "text-align: left; "
        "font-family: Arial, sans-serif; "
        "color: #333333;"
    )

    table_style = (
        "border-collapse: collapse; "
        "width: 100%; "
        "margin: 20px 0; "
        "box-shadow: 0 2px 4px rgba(0,0,0,0.1);"
    )

    # Format order values
    df_formatted = df.copy()
    df_formatted["Order Value"] = df_formatted["Order Value"].apply(
        lambda x: f"{currency} {x:,.2f}" if pd.notnull(x) and x > 0 else "-"
    )

    # Build HTML
    html = f'<table style="{table_style}">'

    # Header
    html += '<thead><tr>'
    for col in df_formatted.columns:
        html += f'<th style="{header_style}">{col}</th>'
    html += '</tr></thead>'

    # Body with alternating colors
    html += '<tbody>'
    for idx, row in df_formatted.iterrows():
        row_style = "background-color: #f9f9f9;" if idx % 2 == 1 else ""
        html += f'<tr style="{row_style}">'
        for col in df_formatted.columns:
            value = row[col]
            # Right-align order value
            align = "right" if col == "Order Value" else "left"
            style = cell_style + f" text-align: {align};"
            html += f'<td style="{style}">{value}</td>'
        html += '</tr>'
    html += '</tbody>'
    html += '</table>'

    return html


# ============================================================================
# MAIN REPORT SENDER
# ============================================================================

def send_report(
    excel_file: Optional[str] = None,
    summary_sheet: Optional[str] = None,
    date_str: Optional[str] = None,
    additional_attachments: Optional[List[str]] = None
) -> bool:
    """
    Send CSFA report via email.

    Args:
        excel_file: Path to Excel file (optional, uses config default)
        summary_sheet: Name of summary sheet (optional, uses config default)
        date_str: Date string for subject (optional, uses today)
        additional_attachments: Additional files to attach

    Returns:
        True if email sent successfully, False otherwise
    """
    try:
        # Validate configuration
        EmailConfig.validate()

        # Use defaults from config
        excel_file = excel_file or EmailConfig.EXCEL_FILE
        summary_sheet = summary_sheet or EmailConfig.SUMMARY_SHEET
        date_str = date_str or datetime.now().strftime("%Y-%m-%d")

        logger.info(f"📊 Preparing to send report: {excel_file}")

        # Check if Excel file exists
        if not os.path.exists(excel_file):
            logger.error(f"❌ Excel file not found: {excel_file}")
            return False

        # Read summary sheet
        logger.info(f"📖 Reading summary from sheet: {summary_sheet}")
        try:
            df_summary = pd.read_excel(excel_file, sheet_name=summary_sheet)
        except Exception as e:
            logger.error(f"❌ Failed to read Excel file: {e}")
            return False

        # Read all individual rep sheets to build customer detail table
        logger.info(f"📖 Reading individual rep sheets for customer details...")
        df_customer_details = _build_customer_details_table(excel_file, df_summary)

        # Generate KPI section
        logger.info("📈 Calculating KPIs...")
        kpi_html = _generate_kpi_section(df_summary)

        # Generate summary table
        logger.info("🎨 Generating summary table...")
        html_generator = HTMLTableGenerator()
        summary_html_table = html_generator.generate(df_summary, format_money=True)

        # Generate customer details table
        logger.info("🎨 Generating customer details table...")
        customer_details_html = _generate_customer_details_table(df_customer_details)

        # Combine sections
        summary_html = f"""
        {kpi_html}
        <h3 style="color: #4F81BD; margin-top: 30px;">Summary by Salesperson</h3>
        {summary_html_table}
        <h3 style="color: #4F81BD; margin-top: 30px;">All Customer Interactions</h3>
        {customer_details_html}
        """

        # Save HTML preview (optional, for debugging)
        if os.getenv("SAVE_HTML_PREVIEW", "").lower() == "true":
            preview_file = "summary_email_preview.html"
            with open(preview_file, "w", encoding="utf-8") as f:
                f.write(summary_html)
            logger.info(f"💾 HTML preview saved: {preview_file}")

        # Collect attachments
        attachments = [excel_file]
        if additional_attachments:
            attachments.extend(additional_attachments)

        # Build email
        logger.info("✉️ Building email message...")
        builder = EmailBuilder(EmailConfig)
        msg = builder.build_message(summary_html, date_str, attachments)

        # Send email
        sender = EmailSender(EmailConfig)
        success = sender.send(msg)

        if success:
            logger.info("🎉 Report sent successfully!")
        else:
            logger.error("❌ Failed to send report")

        return success

    except Exception as e:
        logger.error(f"❌ Error in send_report: {e}", exc_info=True)
        return False


def _generate_kpi_section(df_summary: pd.DataFrame) -> str:
    """Generate KPI summary section with total customers and revenue."""
    try:
        # Get currency from environment or default to MZN
        currency = os.getenv("REPORT_CURRENCY", "MZN")

        # Calculate totals
        total_customers_visited = df_summary["CUSTOMERS VISITED"].sum()
        total_customers_called = df_summary["CUSTOMERS CALLED"].sum()
        total_customers = total_customers_visited + total_customers_called

        total_revenue_visits = df_summary["ORDER VALUE FROM VISITS"].sum()
        total_revenue_calls = df_summary["ORDER VALUE FROM CALLS"].sum()
        total_revenue = total_revenue_visits + total_revenue_calls

        # Format numbers
        total_customers_str = f"{int(total_customers):,}"
        total_revenue_str = f"{currency} {total_revenue:,.2f}"

        # Generate KPI HTML
        kpi_html = f"""
        <div style="margin: 20px 0;">
            <h3 style="color: #4F81BD; margin-bottom: 15px;">Key Performance Indicators</h3>
            <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
                <tr>
                    <td style="padding: 15px; background-color: #E8F4F8; border: 2px solid #4F81BD; width: 50%; text-align: center;">
                        <div style="font-size: 14px; color: #666; margin-bottom: 5px;">TOTAL CUSTOMERS (Visited & Called)</div>
                        <div style="font-size: 28px; font-weight: bold; color: #4F81BD;">{total_customers_str}</div>
                    </td>
                    <td style="padding: 15px; background-color: #E8F4F8; border: 2px solid #4F81BD; width: 50%; text-align: center;">
                        <div style="font-size: 14px; color: #666; margin-bottom: 5px;">TOTAL ORDER REVENUE</div>
                        <div style="font-size: 28px; font-weight: bold; color: #4F81BD;">{total_revenue_str}</div>
                    </td>
                </tr>
            </table>
        </div>
        """

        return kpi_html

    except Exception as e:
        logger.error(f"Error generating KPI section: {e}")
        return ""


# ============================================================================
# STANDALONE EXECUTION
# ============================================================================

def main():
    """Main function for standalone execution."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    logger.info("=" * 60)
    logger.info("CSFA Report Email Sender")
    logger.info("=" * 60)

    success = send_report()

    if success:
        logger.info("✅ Done!")
        return 0
    else:
        logger.error("❌ Failed to send report")
        return 1


if __name__ == "__main__":
    exit(main())
