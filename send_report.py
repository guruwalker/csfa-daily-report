"""
Enhanced Email Module for CSFA Report
Sends daily reports with professional formatting and comprehensive analytics.
"""

import os
import smtplib
import logging
from email.message import EmailMessage
from datetime import datetime
from typing import List, Optional, Dict
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
# DATA ANALYSIS
# ============================================================================

class ReportAnalyzer:
    """Analyze report data for insights."""

    @staticmethod
    def calculate_totals(df: pd.DataFrame) -> Dict:
        """Calculate overall totals."""
        return {
            "total_customers_visited": int(df["CUSTOMERS VISITED"].sum()),
            "total_customers_called": int(df["CUSTOMERS CALLED"].sum()),
            "total_order_value_visits": float(df["ORDER VALUE FROM VISITS"].sum()),
            "total_order_value_calls": float(df["ORDER VALUE FROM CALLS"].sum()),
            "total_salespersons": len(df),
        }

    @staticmethod
    def calculate_averages(df: pd.DataFrame) -> Dict:
        """Calculate average performance metrics."""
        return {
            "avg_customers_visited": df["CUSTOMERS VISITED"].mean(),
            "avg_customers_called": df["CUSTOMERS CALLED"].mean(),
            "avg_order_value_visits": df["ORDER VALUE FROM VISITS"].mean(),
            "avg_order_value_calls": df["ORDER VALUE FROM CALLS"].mean(),
        }


# ============================================================================
# HTML TABLE GENERATOR
# ============================================================================

class HTMLTableGenerator:
    """Generate HTML tables from DataFrames with professional styling."""

    # Enhanced styling constants
    HEADER_STYLE = (
        "background: linear-gradient(135deg, #4F81BD 0%, #2F5F8D 100%); "
        "color: #ffffff; "
        "font-weight: bold; "
        "padding: 14px 12px; "
        "text-align: left; "
        "border: 1px solid #2F5F8D; "
        "font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; "
        "font-size: 13px; "
        "letter-spacing: 0.3px; "
        "text-transform: uppercase;"
    )

    CELL_STYLE = (
        "padding: 12px; "
        "border: 1px solid #e0e0e0; "
        "text-align: left; "
        "font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; "
        "font-size: 13px;"
    )

    TABLE_STYLE = (
        "border-collapse: collapse; "
        "width: 100%; "
        "margin: 25px 0; "
        "box-shadow: 0 4px 6px rgba(0,0,0,0.1); "
        "border-radius: 8px; "
        "overflow: hidden;"
    )

    ALT_ROW_STYLE = "background-color: #f8f9fa;"

    @classmethod
    def generate(cls, df: pd.DataFrame, format_money: bool = True) -> str:
        """
        Generate HTML table from DataFrame with enhanced styling.

        Args:
            df: DataFrame to convert
            format_money: Whether to format numeric columns as money

        Returns:
            HTML table string
        """
        if df.empty:
            return "<p style='font-style: italic; opacity: 0.7;'>No data available</p>"

        # Format numeric columns
        df_formatted = df.copy()
        numeric_cols = df_formatted.select_dtypes(include=['number']).columns

        if format_money:
            for col in numeric_cols:
                df_formatted[col] = df_formatted[col].apply(
                    lambda x: f"{x:,.2f}" if pd.notnull(x) else "0.00"
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
                align = "right" if col in numeric_cols else "left"
                cell_style = cls.CELL_STYLE + f" text-align: {align};"
                html += f'<td style="{cell_style}">{value}</td>'
            html += '</tr>'
        html += '</tbody>'
        html += '</table>'

        return html

    @classmethod
    def generate_metrics_card(cls, title: str, value: str, subtitle: str = "", color: str = "#4F81BD") -> str:
        """Generate a metric card HTML."""
        return f"""
        <div style="
            display: inline-block;
            background: white;
            border-left: 4px solid {color};
            padding: 16px 20px;
            margin: 10px 10px 10px 0;
            border-radius: 6px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
            min-width: 180px;
        ">
            <div style="
                font-size: 12px;
                opacity: 0.7;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                margin-bottom: 8px;
            ">{title}</div>
            <div style="
                font-size: 28px;
                font-weight: bold;
                color: {color};
                margin-bottom: 4px;
            ">{value}</div>
            {f'<div style="font-size: 11px; opacity: 0.6;">{subtitle}</div>' if subtitle else ''}
        </div>
        """


# ============================================================================
# EMAIL BUILDER
# ============================================================================

class EmailBuilder:
    """Build email messages with attachments and professional formatting."""

    def __init__(self, config: EmailConfig):
        self.config = config
        self.html_gen = HTMLTableGenerator()
        self.analyzer = ReportAnalyzer()

    def build_message(
        self,
        df_summary: pd.DataFrame,
        date_str: str,
        attachments: Optional[List[str]] = None
    ) -> EmailMessage:
        """
        Build complete email message with comprehensive analytics.

        Args:
            df_summary: Summary DataFrame
            date_str: Date string for subject
            attachments: List of file paths to attach

        Returns:
            EmailMessage object ready to send
        """
        msg = EmailMessage()

        # Set headers
        msg["From"] = f"{self.config.SENDER_NAME} <{self.config.SENDER_EMAIL}>"
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
        body = self._build_html_body(df_summary, date_str)
        msg.add_alternative(body, subtype="html")

        # Add attachments
        if attachments:
            for filepath in attachments:
                self._attach_file(msg, filepath)

        return msg

    def _build_html_body(self, df_summary: pd.DataFrame, date_str: str) -> str:
        """Build comprehensive HTML email body with analytics."""

        # Calculate analytics
        totals = self.analyzer.calculate_totals(df_summary)
        averages = self.analyzer.calculate_averages(df_summary)

        # Generate summary table
        summary_html = self.html_gen.generate(df_summary, format_money=True)

        # Build metrics cards
        metrics_html = self._build_metrics_section(totals, averages)

        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body {{
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
                    line-height: 1.6;
                    background-color: #f5f5f5;
                    margin: 0;
                    padding: 0;
                }}
                .container {{
                    max-width: 900px;
                    margin: 0 auto;
                    background-color: #ffffff;
                    padding: 0;
                }}
                .header {{
                    background: linear-gradient(135deg, #4F81BD 0%, #2F5F8D 100%);
                    color: #ffffff;
                    padding: 30px;
                    text-align: center;
                }}
                .header h1 {{
                    margin: 0;
                    font-size: 26px;
                    font-weight: 600;
                    color: #ffffff;
                }}
                .header p {{
                    margin: 10px 0 0 0;
                    font-size: 14px;
                    opacity: 0.9;
                    color: #ffffff;
                }}
                .content {{
                    padding: 30px;
                }}
                .section {{
                    margin-bottom: 35px;
                }}
                .section-title {{
                    font-size: 20px;
                    font-weight: 600;
                    color: #2F5F8D;
                    margin-bottom: 20px;
                    padding-bottom: 10px;
                    border-bottom: 2px solid #4F81BD;
                }}
                .metrics-grid {{
                    display: flex;
                    flex-wrap: wrap;
                    gap: 15px;
                    margin: 20px 0;
                }}
                .info-box {{
                    background: #f8f9fa;
                    border-left: 4px solid #4F81BD;
                    padding: 15px 20px;
                    margin: 15px 0;
                    border-radius: 4px;
                }}
                .footer {{
                    background-color: #f8f9fa;
                    padding: 25px 30px;
                    margin-top: 30px;
                    border-top: 3px solid #4F81BD;
                }}
                .footer-signature {{
                    margin: 15px 0;
                    line-height: 1.8;
                }}
                .footer-disclaimer {{
                    margin-top: 20px;
                    padding-top: 15px;
                    border-top: 1px solid #ddd;
                    font-size: 11px;
                    opacity: 0.7;
                    font-style: italic;
                }}

                /* Dark mode support */
                @media (prefers-color-scheme: dark) {{
                    body {{
                        background-color: #1a1a1a;
                    }}
                    .container {{
                        background-color: #2d2d2d;
                        color: #e0e0e0;
                    }}
                    .section-title {{
                        color: #6fa3d8;
                    }}
                    .info-box {{
                        background: #3a3a3a;
                        color: #e0e0e0;
                    }}
                    .footer {{
                        background-color: #3a3a3a;
                        color: #e0e0e0;
                    }}
                }}

                /* Mobile responsiveness */
                @media only screen and (max-width: 600px) {{
                    .content {{
                        padding: 20px;
                    }}
                    .header h1 {{
                        font-size: 22px;
                    }}
                    .section-title {{
                        font-size: 18px;
                    }}
                    .metrics-grid {{
                        flex-direction: column;
                    }}
                    table {{
                        font-size: 12px;
                    }}
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📊 Tintas Berger CSFA Daily Report</h1>
                    <p>{date_str} | Customer Sales & Field Activity</p>
                </div>

                <div class="content">
                    <div class="section">
                        <p style="font-size: 15px;">Dear {self.config.RECIPIENT_NAME},</p>
                        <p style="font-size: 14px; line-height: 1.8;">
                            Please find the daily Customer Sales and Field Activity (CSFA) report for <strong>{date_str}</strong>.
                        </p>
                    </div>

                    <div class="section">
                        <div class="section-title">📈 Key Performance Indicators</div>
                        {metrics_html}
                    </div>

                    <div class="section">
                        <div class="section-title">👥 Detailed Performance by Salesperson</div>
                        {summary_html}
                    </div>

                    <div class="info-box">
                        <strong>📎 Attachment:</strong> The complete detailed Excel report is attached.
                    </div>
                </div>

                <div class="footer">
                    <div class="footer-signature">
                        Kind regards,<br>
                        <strong>{self.config.SENDER_NAME}</strong>
                    </div>

                    <div class="footer-disclaimer">
                        This is an automated report. Please do not reply to this email.
                    </div>
                </div>
            </div>
        </body>
        </html>
        """

    def _build_metrics_section(self, totals: Dict, averages: Dict) -> str:
        """Build metrics cards section."""
        total_customers = totals["total_customers_visited"] + totals["total_customers_called"]
        total_revenue = totals["total_order_value_visits"] + totals["total_order_value_calls"]

        metrics = [
            self.html_gen.generate_metrics_card(
                "Total Customers",
                f"{total_customers:,}",
                f"{totals['total_customers_visited']} visited, {totals['total_customers_called']} called",
                "#4F81BD"
            ),
            self.html_gen.generate_metrics_card(
                "Total Revenue",
                f"{total_revenue:,.2f} MZN",
                f"From {totals['total_salespersons']} salespersons",
                "#28a745"
            ),
            self.html_gen.generate_metrics_card(
                "Avg. Customers/Rep",
                f"{total_customers / totals['total_salespersons']:.1f}",
                "Per salesperson",
                "#ff6b6b"
            ),
            self.html_gen.generate_metrics_card(
                "Avg. Order Value",
                f"{total_revenue / total_customers:,.2f} MZN",
                "Per customer",
                "#ffa500"
            ),
        ]

        return '<div class="metrics-grid">' + ''.join(metrics) + '</div>'

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
# MAIN REPORT SENDER
# ============================================================================

def send_report(
    excel_file: Optional[str] = None,
    summary_sheet: Optional[str] = None,
    date_str: Optional[str] = None,
    additional_attachments: Optional[List[str]] = None
) -> bool:
    """
    Send CSFA report via email with comprehensive analytics.

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

        # Save HTML preview (optional, for debugging)
        if os.getenv("SAVE_HTML_PREVIEW", "").lower() == "true":
            preview_file = "email_preview.html"
            builder = EmailBuilder(EmailConfig)
            html_content = builder._build_html_body(df_summary, date_str)
            with open(preview_file, "w", encoding="utf-8") as f:
                f.write(html_content)
            logger.info(f"💾 HTML preview saved: {preview_file}")

        # Collect attachments
        attachments = [excel_file]
        if additional_attachments:
            attachments.extend(additional_attachments)

        # Build email
        logger.info("✉️ Building email message...")
        builder = EmailBuilder(EmailConfig)
        msg = builder.build_message(df_summary, date_str, attachments)

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
