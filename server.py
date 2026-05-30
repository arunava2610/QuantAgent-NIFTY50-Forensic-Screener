import os
import time
import requests
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import docx
from docx.shared import Pt, RGBColor, Inches
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from google import genai as g_genai
from dotenv import load_dotenv

# ---------------------------------------------------------
# CONFIGURATION & LOGGING
# ---------------------------------------------------------
load_dotenv()

# Configure terminal logger
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

app = FastAPI(title="QuantAgent API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class ReportRequest(BaseModel):
    companies: list[str]
    custom_feedback: str
    email: str = "" 

# ⚠️ INSERT YOUR GOOGLE DRIVE FILE IDs HERE
EXCEL_DRIVE_ID = "1WQljak5wURGcg5izR_jPFfMwcqza97jx"
API_KEY_DRIVE_ID = "1eaOBvwclf6xtIxfT3RnJdE3dwCKm1b_5"
EXCEL_DRIVE_URL = f"https://drive.google.com/uc?export=download&id={EXCEL_DRIVE_ID}"
API_KEY_DRIVE_URL = f"https://drive.google.com/uc?export=download&id={API_KEY_DRIVE_ID}"

cached_companies = []
master_financial_df = None

# ---------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------
def get_api_key():
    try:
        response = requests.get(API_KEY_DRIVE_URL)
        if response.status_code == 200:
            return response.text.strip()
    except Exception as e:
        logger.error(f"Failed to fetch API key from Drive: {e}")
    return None

def download_excel_from_drive():
    """Safely downloads the Excel file from Google Drive to bypass Pandas URL blocking."""
    temp_file_path = "temp_drive_matrix.xlsx"
    try:
        # We use a standard browser User-Agent so Google Drive doesn't block the request
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        response = requests.get(EXCEL_DRIVE_URL, headers=headers)
        
        if response.status_code == 200:
            with open(temp_file_path, "wb") as f:
                f.write(response.content)
            return temp_file_path
        else:
            logger.error(f"Google Drive returned status code: {response.status_code}")
            return None
    except Exception as e:
        logger.error(f"Failed to download Excel from Drive: {e}")
        return None

def send_report_via_email(recipient_email, file_path):
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_PASSWORD")
    
    if not sender_email or not sender_password:
        raise Exception("SMTP credentials not configured in backend environment.")

    logger.info(f"Preparing to send report to {recipient_email} via SMTP...")

    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "NIFTY Investment Committee Executive Screening Document"

        body = (
            "Hello,\n\nPlease find attached the finalized executive screening report "
            "covering your selected enterprise data profiles.\n\nAll fundamental metrics and debt ratios "
            "were processed securely and verified against our automated AI expert engines.\n\n"
            "Best Regards,\nAI Transformation Pipeline Framework"
        )
        msg.attach(MIMEText(body, 'plain'))

        with open(file_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(file_path)}")
            msg.attach(part)

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        logger.info(f"✅ SUCCESS: Email successfully delivered to {recipient_email}")
    except Exception as e:
        logger.error(f"❌ ERROR: Failed to send email to {recipient_email}. Details: {str(e)}")
        raise e

# ---------------------------------------------------------
# API ENDPOINTS
# ---------------------------------------------------------
@app.get("/api/companies")
def get_companies():
    global cached_companies
    if cached_companies:
        return cached_companies

    try:
        logger.info("Fetching fresh companies list from Google Drive...")
        
        # NEW DOWNLOAD LOGIC
        local_excel_path = download_excel_from_drive()
        if not local_excel_path:
            raise Exception("Could not download the Excel file.")

        df = pd.read_excel(local_excel_path, sheet_name="Ratios")
        cached_companies = [{"ticker": str(row["Ticker"]), "name": str(row["Company Name"])} for idx, row in df.iterrows()]
        logger.info(f"Successfully loaded {len(cached_companies)} companies into memory.")
        return cached_companies
    except Exception as e:
        logger.error(f"Drive fetch failed: {e}")
        return [{"ticker": "RELIANCE.NS", "name": "Reliance Industries"}]


@app.post("/api/generate-report")
def generate_report(request: ReportRequest):
    global master_financial_df

    logger.info(f"Incoming report request for {len(request.companies)} companies.")
    if request.email:
        logger.info(f"Target Delivery Method: Email -> {request.email}")
    else:
        logger.info("Target Delivery Method: Local Download")

    live_api_key = get_api_key()
    if not live_api_key:
        return {"error": "Failed to retrieve Gemini API Key from Google Drive."}
    
    client = g_genai.Client(api_key=live_api_key)

    if master_financial_df is None:
        logger.info("Downloading massive historical matrix from Drive...")
        
        # NEW DOWNLOAD LOGIC
        local_excel_path = download_excel_from_drive()
        if not local_excel_path:
             return {"error": "Failed to download the financial matrix from Google Drive."}

        xls = pd.ExcelFile(local_excel_path)
        all_years_data = []
        for sheet_name in xls.sheet_names:
            if sheet_name != "Ratios":
                year_df = pd.read_excel(xls, sheet_name=sheet_name, header=2)
                year_df['Fiscal_Year'] = int(sheet_name)
                all_years_data.append(year_df)
        master_financial_df = pd.concat(all_years_data, ignore_index=True)
        logger.info("Historical matrix successfully cached.")

    compiled_prompt = "You are an Elite Corporate Investment Committee Member. Evaluate the following securities.\n\n"
    company_metadata = [] 

    for ticker in request.companies:
        comp_annual_df = master_financial_df[(master_financial_df['Ticker'] == ticker) & (master_financial_df['Type'] == 'Annual')].copy()
        if comp_annual_df.empty:
            continue
            
        company_name = comp_annual_df['Company Name'].iloc[0]
        comp_annual_df = comp_annual_df.sort_values('Fiscal_Year')

        comp_annual_df['Net Profit Margin (%)'] = (comp_annual_df['Net Income'] / comp_annual_df['Total Revenue']) * 100
        comp_annual_df['Return on Assets (%)'] = (comp_annual_df['Net Income'] / comp_annual_df['Total Assets']) * 100
        comp_annual_df['Debt to Assets (%)'] = (comp_annual_df['Long Term Debt'].fillna(0) / comp_annual_df['Total Assets']) * 100
        comp_annual_df['Cash to Assets (%)'] = (comp_annual_df['Cash And Cash Equivalents'].fillna(0) / comp_annual_df['Total Assets']) * 100
        
        ratios_summary_text = ""
        valid_ratio_df = comp_annual_df.dropna(subset=['Net Profit Margin (%)', 'Total Revenue'])
        for _, r_row in valid_ratio_df.iterrows():
            ratios_summary_text += f"• FY {int(r_row['Fiscal_Year'])} | NPM: {r_row['Net Profit Margin (%)']:.2f}% | ROA: {r_row['Return on Assets (%)']:.2f}% | Debt/Assets: {r_row['Debt to Assets (%)']:.2f}%\n"

        comp_all_df = master_financial_df[master_financial_df['Ticker'] == ticker].dropna(how='all', axis=1)
        compressed_ledger_csv = comp_all_df.to_csv(index=False)

        company_metadata.append({
            "ticker": ticker,
            "name": company_name,
            "valid_ratio_df": valid_ratio_df
        })

        compiled_prompt += f"--- TARGET SECURITY: {ticker} ({company_name}) ---\n"
        compiled_prompt += f"[HISTORICAL LEDGER]\n{compressed_ledger_csv}\n"
        compiled_prompt += f"[RATIOS PROFILE]\n{ratios_summary_text}\n\n"

    if request.custom_feedback:
        compiled_prompt += f"[USER STRATEGIC DIRECTIVES]\nPay special attention to these directives: {request.custom_feedback}\n\n"

    compiled_prompt += """
    Output STRICTLY in this exact format. Do not add conversational text. You MUST include every requested ticker:
    TICKER: [Exact Ticker Symbol]
    RATING: [BUY, SELL, or HOLD]
    REASONING: [Dense, high-impact 4-5 sentence rationale citing specific numbers.]
    ===
    """

    # ---------------------------------------------------------
    # GEMINI API CALL WITH RETRY LOGIC
    # ---------------------------------------------------------
    max_retries = 3
    retry_delay = 5
    ai_response = ""

    for attempt in range(1, max_retries + 1):
        try:
            logger.info(f"Sending batched payload to Gemini (Attempt {attempt}/{max_retries})...")
            response = client.models.generate_content(model='gemini-2.5-flash', contents=compiled_prompt)
            ai_response = response.text
            logger.info("✅ Gemini API generation successful.")
            break  # Exit loop on success
        except Exception as e:
            logger.warning(f"⚠️ Gemini API failed on attempt {attempt}. Error: {str(e)}")
            if attempt < max_retries:
                logger.info(f"Waiting {retry_delay} seconds before retrying...")
                time.sleep(retry_delay)
            else:
                logger.error("❌ Max retries reached. Gemini API failed completely.")
                ai_response = ""

    # Parse AI response
    ai_results = {}
    current_ticker = None
    current_rating = "REVIEW REQUIRED"
    current_reasoning = ""

    for line in ai_response.split('\n'):
        clean_line = line.strip()
        upper_line = clean_line.upper()
        
        if upper_line.startswith("TICKER:"):
            if current_ticker:
                ai_results[current_ticker] = {"rating": current_rating, "reasoning": current_reasoning.strip()}
            current_ticker = clean_line.split(":", 1)[1].strip()
            current_reasoning = ""
            current_rating = "REVIEW REQUIRED"
        elif upper_line.startswith("RATING:"):
            current_rating = clean_line.split(":", 1)[1].replace("*", "").strip()
        elif upper_line.startswith("REASONING:"):
            current_reasoning = clean_line.split(":", 1)[1].strip() + " "
        elif upper_line == "===" or upper_line.startswith("---"):
            pass
        elif current_ticker and clean_line:
            current_reasoning += clean_line + " "

    if current_ticker:
        ai_results[current_ticker] = {"rating": current_rating, "reasoning": current_reasoning.strip()}

    # Document generation
    logger.info("Compiling Word Document and generating charts...")
    doc = docx.Document()
    doc.add_heading('NIFTY Executive Investment Screening Summary', level=0)
    doc.add_paragraph('Generated by: Hybrid Agentic AI & Custom Strategy Directives')
    
    if request.custom_feedback:
        doc.add_heading('User Strategic Guidelines Applied:', level=2)
        doc.add_paragraph(request.custom_feedback, style='Intense Quote')

    for idx, meta in enumerate(company_metadata):
        ticker = meta["ticker"]
        company_name = meta["name"]
        valid_ratio_df = meta["valid_ratio_df"]

        result = ai_results.get(ticker, {"rating": "HOLD", "reasoning": "AI Generation failed or skipped this asset due to API limits/errors."})
        
        chart_filename = f"temp_chart_{idx}.png"
        try:
            fig, ax = plt.subplots(figsize=(7, 4))
            years_str = valid_ratio_df['Fiscal_Year'].astype(str)
            ax.plot(years_str, valid_ratio_df['Net Profit Margin (%)'], marker='o', linewidth=2.5, label='NPM %')
            ax.plot(years_str, valid_ratio_df['Return on Assets (%)'], marker='s', linewidth=2.5, label='ROA %')
            ax.plot(years_str, valid_ratio_df['Debt to Assets (%)'], marker='^', linewidth=2.5, label='Debt/Assets %')
            
            ax.set_title(f"Financial Ratios Profile: {company_name}", fontsize=11, fontweight='bold')
            ax.legend(loc='best', fontsize=8)
            ax.grid(True, linestyle='--', alpha=0.6)
            plt.tight_layout()
            plt.savefig(chart_filename, dpi=150)
            plt.close(fig)
        except Exception as e:
            logger.error(f"Chart generation error for {ticker}: {e}")

        p_heading = doc.add_paragraph()
        p_heading.add_run(f"\n🏢 {company_name}").bold = True
        p_heading.add_run(" | Action: ").italic = True
        r_rate = p_heading.add_run(result["rating"])
        r_rate.bold = True
        
        if "BUY" in result["rating"].upper(): r_rate.font.color.rgb = RGBColor(0, 128, 0)
        elif "SELL" in result["rating"].upper(): r_rate.font.color.rgb = RGBColor(180, 0, 0)

        doc.add_paragraph("🧠 Forensic Rationale:").bold = True
        doc.add_paragraph(result["reasoning"])

        if os.path.exists(chart_filename):
            doc.add_paragraph().add_run().add_picture(chart_filename, width=Inches(5.5))
            os.remove(chart_filename) 

    output_path = "NIFTY_Custom_Screening_Report.docx"
    doc.save(output_path)
    logger.info("Document successfully saved to disk.")
    
    # ---------------------------------------------------------
    # ROUTING LOGIC
    # ---------------------------------------------------------
    if request.email and "@" in request.email:
        try:
            send_report_via_email(request.email, output_path)
            os.remove(output_path) 
            return {"message": f"Report successfully generated and dispatched to {request.email}"}
        except Exception as e:
            # If email fails, we still want to inform the frontend
            return {"error": f"Report generated, but failed to send email: {e}"}
    else:
        logger.info("Streaming file back to frontend for local download.")
        return FileResponse(
            path=output_path, 
            filename="NIFTY_Custom_Screening_Report.docx", 
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )