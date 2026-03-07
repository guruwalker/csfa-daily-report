"""
Enhanced Email Module for CSFA Report
Sends daily reports with professional formatting and error handling.
UPDATED: Focus on customer visits, not revenue
UPDATED: Added attendance tracking tables
UPDATED: Simplified holiday email format
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
import re

from holiday_checker import is_holiday

load_dotenv()

logger = logging.getLogger(__name__)


# ============================================================================
# CLASSPROPERTY HELPER
# ============================================================================

class classproperty:
    """
    Descriptor for read-only class-level properties.
    Allows @classproperty on classmethods so os.getenv() is called at
    access time rather than at class definition time.
    """
    def __init__(self, func):
        self.func = func

    def __get__(self, obj, owner):
        return self.func(owner)


# ============================================================================
# EMAIL CONFIGURATION
# ============================================================================

class EmailConfig:
    """
    Email configuration from environment variables.

    All properties are read lazily (at access time) so that changes made to
    os.environ by the Flask app before calling send_report() are picked up
    correctly. This is what makes test vs. live mode switching work.
    """

    # ---- SMTP Settings (static - these don't change between modes) ----
    SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.robbialac.co.mz")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SENDER_EMAIL = os.getenv("SENDER_EMAIL", "innocent.maina@robbialac.co.mz")
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
    SMTP_TIMEOUT = int(os.getenv("SMTP_TIMEOUT", "30"))

    # ---- Static content settings ----
    EMAIL_SUBJECT_TEMPLATE = os.getenv("EMAIL_SUBJECT", "Tintas Berger CSFA Report - {date}")
    SENDER_NAME = os.getenv("SENDER_NAME", "Innocent Maina")
    RECIPIENT_NAME = os.getenv("RECIPIENT_NAME", "Mr. Hussein")
    EXCEL_FILE = os.getenv("OUTPUT_FILE", "Daily_CSFA_Report.xlsx")
    SUMMARY_SHEET = os.getenv("SUMMARY_SHEET", "Day Summary")
    INCLUDE_MONTHLY_ATTENDANCE = os.getenv("INCLUDE_MONTHLY_ATTENDANCE", "true").lower() == "true"

    # ---- Recipients and threading: READ LAZILY via properties ----
    # These must be properties (not class-level assignments) so that when
    # app.py sets os.environ["EMAIL_TO"] / os.environ["EMAIL_THREAD_ID"]
    # before calling send_report(), the new values are actually used.

    @classproperty
    def TO_RECIPIENTS(cls) -> List[str]:
        return os.getenv("EMAIL_TO", "innocent.maina@crownpaints.co.ke").split(",")

    @classproperty
    def CC_RECIPIENTS(cls) -> List[str]:
        raw = os.getenv("EMAIL_CC", "")
        return raw.split(",") if raw else []

    @classproperty
    def BCC_RECIPIENTS(cls) -> List[str]:
        raw = os.getenv("EMAIL_BCC", "")
        return raw.split(",") if raw else []

    @classproperty
    def EMAIL_THREAD_ID(cls) -> Optional[str]:
        return os.getenv("EMAIL_THREAD_ID") or None

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
    def generate(cls, df: pd.DataFrame, format_money: bool = False) -> str:
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
        customer_columns = ['CUSTOMERS VISITED']
        for col in customer_columns:
            if col in df_formatted.columns:
                df_formatted[col] = df_formatted[col].apply(
                    lambda x: f"{int(x):,}" if pd.notnull(x) else "0"
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
                is_numeric = col in customer_columns
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

        # ========== IMPROVED THREADING LOGIC ==========
        import time

        thread_id = self.config.EMAIL_THREAD_ID
        base_subject = self.config.EMAIL_SUBJECT_TEMPLATE

        # Set subject - keep it consistent for threading
        if thread_id:
            # Reply - add Re: prefix and keep base subject
            subject = f"Re: {base_subject}"
        else:
            # First email
            subject = base_subject

        msg["Subject"] = subject

        # Handle Message-ID and threading headers
        if not thread_id:
            # First email - generate a Message-ID
            timestamp = str(int(time.time() * 1000))
            hostname = self.config.SENDER_EMAIL.split("@")[1]
            message_id = f"<csfa-report-{timestamp}@{hostname}>"
            msg["Message-ID"] = message_id
            logger.info(f"📧 Generated new Message-ID: {message_id}")
            logger.info(f"⚠️  SAVE THIS to .env as EMAIL_THREAD_ID for threading!")
        else:
            # Reply to existing thread
            # In-Reply-To should be the ORIGINAL message ID (first in thread)
            msg["In-Reply-To"] = thread_id

            # References should be the SAME as In-Reply-To for consistent threading
            # Email clients will group messages with the same References header
            msg["References"] = thread_id

            logger.info(f"📧 Threading email to: {thread_id}")
        # ==============================================

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
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
            day_name = date_obj.strftime("%A")
            formatted_date = f"{day_name}, {date_str}"
        except:
            formatted_date = date_str

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
# KPI SECTION GENERATOR
# ============================================================================

def _generate_kpi_section(df_summary: pd.DataFrame) -> str:
    """Generate KPI summary section focused on customer visits."""
    try:
        # Calculate totals
        total_customers_visited = int(df_summary["CUSTOMERS VISITED"].sum())

        # Calculate active salespeople (those who used the app)
        active_salespeople = int(df_summary[df_summary["APP USAGE"] == "Used App"].shape[0])
        total_salespeople = len(df_summary)
        active_fraction = f"{active_salespeople}/{total_salespeople}"

        # Format numbers
        visited_str = f"{total_customers_visited:,}"

        # Generate KPI HTML (2-column layout)
        kpi_html = f"""
        <div style="margin: 20px 0;">
            <h3 style="color: #4F81BD; margin-bottom: 15px;">Key Performance Indicators</h3>
            <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
                <tr>
                    <td style="padding: 15px; background-color: #E8F4F8; border: 2px solid #4F81BD; width: 50%; text-align: center;">
                        <div style="font-size: 14px; color: #666; margin-bottom: 5px;">CUSTOMERS VISITED</div>
                        <div style="font-size: 28px; font-weight: bold; color: #4F81BD;">{visited_str}</div>
                    </td>
                    <td style="padding: 15px; background-color: #E8F4F8; border: 2px solid #4F81BD; width: 50%; text-align: center;">
                        <div style="font-size: 14px; color: #666; margin-bottom: 5px;">ACTIVE SALESPEOPLE</div>
                        <div style="font-size: 28px; font-weight: bold; color: #4F81BD;">{active_fraction}</div>
                    </td>
                </tr>
            </table>
        </div>
        """

        return kpi_html

    except Exception as e:
        logger.error(f"Error generating KPI section: {e}")
        return ""


def _get_attendance_grid_sheet_name(excel_file: str) -> str:
    """
    Find the Attendance Grid sheet name in the Excel file.
    Named like "Attendance Grid - February".

    Args:
        excel_file: Path to Excel file

    Returns:
        Sheet name if found, None otherwise
    """
    from openpyxl import load_workbook

    try:
        wb = load_workbook(excel_file, read_only=True)
        for sheet_name in wb.sheetnames:
            if sheet_name.startswith("Attendance Grid -"):
                wb.close()
                return sheet_name
        wb.close()
    except Exception as e:
        logger.warning(f"Could not search for attendance grid sheet: {e}")

    return None


def _generate_attendance_grid_html(df: pd.DataFrame, month_name: str) -> str:
    """
    Render the Attendance Grid DataFrame as an HTML table for email.

    The grid has columns: Salesperson | 1 | 2 | 3 | ... | Summary
    Cells contain ✓ (present), X (absent), L (leave), or "" (future/weekend).

    Color coding matches the Excel sheet:
      ✓  → light green
      X  → light red
      L  → light yellow
      "" → no fill (future days)
    """
    if df is None or df.empty:
        return ""

    CELL_COLORS = {
        "✓": "#C6EFCE",   # light green
        "X": "#FFC7CE",   # light red / absent
        "L": "#FFEB9C",   # light yellow / leave
    }

    TABLE_STYLE = (
        "border-collapse: collapse; "
        "width: 100%; "
        "margin: 10px 0 20px 0; "
        "font-family: Arial, sans-serif; "
        "font-size: 12px;"
    )
    HEADER_STYLE = (
        "background-color: #4F81BD; "
        "color: #FFFFFF; "
        "font-weight: bold; "
        "padding: 6px 4px; "
        "text-align: center; "
        "border: 1px solid #2F5F8D; "
        "white-space: nowrap;"
    )
    NAME_HEADER_STYLE = (
        "background-color: #4F81BD; "
        "color: #FFFFFF; "
        "font-weight: bold; "
        "padding: 6px 8px; "
        "text-align: left; "
        "border: 1px solid #2F5F8D;"
    )
    NAME_CELL_STYLE = (
        "padding: 5px 8px; "
        "border: 1px solid #ddd; "
        "text-align: left; "
        "color: #333; "
        "white-space: nowrap;"
    )
    DAY_CELL_BASE = (
        "padding: 5px 4px; "
        "border: 1px solid #ddd; "
        "text-align: center; "
        "min-width: 18px;"
    )
    SUMMARY_CELL_STYLE = (
        "padding: 5px 8px; "
        "border: 1px solid #ddd; "
        "text-align: left; "
        "color: #333; "
        "white-space: nowrap;"
    )

    columns = list(df.columns)  # ["Salesperson", "1", "2", ..., "Summary"]

    html = f'<table style="{TABLE_STYLE}">'

    # Header row
    html += "<thead><tr>"
    for col in columns:
        if col == "Salesperson":
            html += f'<th style="{NAME_HEADER_STYLE}">{col}</th>'
        elif col == "Summary":
            html += f'<th style="{HEADER_STYLE} text-align: left;">{col}</th>'
        else:
            html += f'<th style="{HEADER_STYLE}">{col}</th>'
    html += "</tr></thead>"

    # Data rows
    html += "<tbody>"
    for row_idx, (_, row) in enumerate(df.iterrows()):
        bg = "#f9f9f9" if row_idx % 2 == 1 else "#ffffff"
        html += f'<tr style="background-color: {bg};">'

        for col in columns:
            val = str(row[col]) if pd.notnull(row[col]) else ""

            if col == "Salesperson":
                html += f'<td style="{NAME_CELL_STYLE}">{val}</td>'
            elif col == "Summary":
                html += f'<td style="{SUMMARY_CELL_STYLE}">{val}</td>'
            else:
                # Day cell — apply colour based on value
                cell_color = CELL_COLORS.get(val, bg)
                html += (
                    f'<td style="{DAY_CELL_BASE} background-color: {cell_color};">'
                    f'{val}</td>'
                )

        html += "</tr>"
    html += "</tbody></table>"

    # Legend
    legend_items = [
        ("#C6EFCE", "✓ Present"),
        ("#FFC7CE", "X Absent"),
        ("#FFEB9C", "L Leave"),
        ("#ffffff", "  Future"),
    ]
    legend_html = '<div style="font-family: Arial, sans-serif; font-size: 11px; margin-bottom: 10px;">'
    legend_html += '<strong>Legend:</strong>&nbsp;&nbsp;'
    for color, label in legend_items:
        legend_html += (
            f'<span style="display:inline-block; background:{color}; '
            f'border:1px solid #ccc; padding: 2px 8px; margin-right:8px;">{label}</span>'
        )
    legend_html += '</div>'

    return (
        f'<h3 style="color: #4F81BD; margin-top: 30px;">'
        f'Monthly Attendance Grid - {month_name}</h3>'
        + legend_html
        + html
    )


def _send_holiday_report(excel_file: str, date_str: str) -> bool:
    """
    Send simplified holiday report email.

    Args:
        excel_file: Path to Excel file
        date_str: Date string

    Returns:
        True if successful
    """
    try:
        # Parse date
        report_date = datetime.strptime(date_str, "%Y-%m-%d")
        day_name = report_date.strftime("%A")
        formatted_date = f"{day_name}, {date_str}"

        # Build simple holiday email (no unnecessary boilerplate)
        holiday_body = f"""
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
            </style>
        </head>
        <body>
            <p>Greetings {EmailConfig.RECIPIENT_NAME},</p>

            <p>{formatted_date} was a public holiday, no CSFA activity recorded.</p>

            <p>Kind regards,<br>
            <strong>{EmailConfig.SENDER_NAME}</strong></p>
        </body>
        </html>
        """

        # Build email message directly (bypass the normal builder to avoid boilerplate)
        msg = EmailMessage()

        # Set headers
        msg["From"] = EmailConfig.SENDER_EMAIL
        msg["To"] = ", ".join(EmailConfig.clean_recipients(EmailConfig.TO_RECIPIENTS))

        cc_recipients = EmailConfig.clean_recipients(EmailConfig.CC_RECIPIENTS)
        if cc_recipients:
            msg["Cc"] = ", ".join(cc_recipients)

        bcc_recipients = EmailConfig.clean_recipients(EmailConfig.BCC_RECIPIENTS)
        if bcc_recipients:
            msg["Bcc"] = ", ".join(bcc_recipients)

        # Subject with [Holiday] prefix for easy identification
        import time
        thread_id = EmailConfig.EMAIL_THREAD_ID
        base_subject = EmailConfig.EMAIL_SUBJECT_TEMPLATE.replace("{date}", date_str)

        if thread_id:
            subject = f"Re: {base_subject} [Holiday]"
        else:
            subject = f"{base_subject} [Holiday]"

        msg["Subject"] = subject

        # Threading headers
        if not thread_id:
            timestamp = str(int(time.time() * 1000))
            hostname = EmailConfig.SENDER_EMAIL.split("@")[1]
            message_id = f"<csfa-report-{timestamp}@{hostname}>"
            msg["Message-ID"] = message_id
            logger.info(f"📧 Generated new Message-ID: {message_id}")
        else:
            msg["In-Reply-To"] = thread_id
            msg["References"] = thread_id
            logger.info(f"📧 Threading holiday email to: {thread_id}")

        # Add HTML body
        msg.add_alternative(holiday_body, subtype="html")

        # Attach Excel file if it exists
        if os.path.exists(excel_file):
            with open(excel_file, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype="application",
                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename=os.path.basename(excel_file)
                )
            logger.info(f"✅ Attached: {os.path.basename(excel_file)}")

        # Send email
        sender = EmailSender(EmailConfig)
        success = sender.send(msg)

        if success:
            logger.info("🎉 Holiday report email sent successfully!")

        return success

    except Exception as e:
        logger.error(f"❌ Error sending holiday report: {e}", exc_info=True)
        return False


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
    Send CSFA report via email with monthly attendance tracking.

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

        # Check if this is a holiday report
        try:
            report_date = datetime.strptime(date_str, "%Y-%m-%d")
            if is_holiday(report_date):
                logger.info("🎉 Holiday detected - sending simplified holiday email")
                return _send_holiday_report(excel_file, date_str)
        except ValueError:
            pass  # If date parsing fails, continue with normal report

        # Normal report processing continues here...
        # Read summary sheet
        logger.info(f"📖 Reading summary from sheet: {summary_sheet}")
        try:
            df_summary = pd.read_excel(excel_file, sheet_name=summary_sheet)
        except Exception as e:
            logger.error(f"❌ Failed to read Excel file: {e}")
            return False

        # Read Attendance Grid sheet if enabled
        df_attendance_grid = None
        attendance_grid_sheet_name = None
        month_name = None

        if EmailConfig.INCLUDE_MONTHLY_ATTENDANCE:
            try:
                attendance_grid_sheet_name = _get_attendance_grid_sheet_name(excel_file)

                if attendance_grid_sheet_name:
                    logger.info(f"📖 Reading attendance grid from sheet: {attendance_grid_sheet_name}")
                    df_attendance_grid = pd.read_excel(excel_file, sheet_name=attendance_grid_sheet_name)
                    month_name = attendance_grid_sheet_name.split(" - ")[-1] if attendance_grid_sheet_name else "This Month"
                else:
                    logger.warning("⚠️ Attendance Grid sheet not found")
            except Exception as e:
                logger.warning(f"⚠️ Could not read attendance grid sheet: {e}")
                logger.info("Continuing without attendance data in email...")

        # Generate KPI sections
        logger.info("📈 Calculating KPIs...")
        kpi_html = _generate_kpi_section(df_summary)

        # Generate summary table
        logger.info("🎨 Generating summary table...")
        html_generator = HTMLTableGenerator()
        summary_html_table = html_generator.generate(df_summary, format_money=False)

        # Generate attendance grid HTML if available
        monthly_attendance_html = ""
        if df_attendance_grid is not None and not df_attendance_grid.empty:
            logger.info("🎨 Generating attendance grid table...")
            monthly_attendance_html = _generate_attendance_grid_html(df_attendance_grid, month_name)
        elif EmailConfig.INCLUDE_MONTHLY_ATTENDANCE:
            monthly_attendance_html = (
                f'<h3 style="color: #4F81BD; margin-top: 30px;">'
                f'Monthly Attendance Grid - {month_name or "This Month"}</h3>'
                f'<p style="color: #666; font-style: italic;">'
                f'No attendance recorded yet for {month_name or "this month"}.</p>'
            )

        # Combine sections (removed customer details table)
        summary_html = f"""
        {kpi_html}
        <h3 style="color: #4F81BD; margin-top: 30px;">Summary by Salesperson</h3>
        {summary_html_table}
        {monthly_attendance_html}
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
    logger.info("CSFA Report Email Sender (with Monthly Attendance)")
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
