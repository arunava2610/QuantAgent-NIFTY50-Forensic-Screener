import sys
import os
import smtplib
import time  # Added for retry delay logic
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google import genai as g_genai

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import logger, get_api_key

def generate_executive_summary(missing_item: str, allocated_node: str, assigned_fleet: str, target_email: str):
    """Drafts a detailed summary and emails it directly to the stakeholder."""
    api_key = get_api_key()
    if not api_key:
        raise Exception("API Key missing.")

    client = g_genai.Client(api_key=api_key) # type: ignore

    prompt = f"""
    You are an executive Supply Chain AI communicating with human leadership. 
    Write a highly detailed, 10-12 line executive summary regarding a recent exception resolution.
    
    You MUST cover these three sections clearly:
    1. The Issue: Explain that a critical failure occurred regarding '{missing_item}'.
    2. Alternatives Considered: Discuss theoretical alternatives (e.g., standard recovery, air freight) and why they were rejected due to SLA risks or cost.
    3. Final Solution: State that inventory was successfully secured from '{allocated_node}' and dispatched via '{assigned_fleet}'.
    
    Maintain a highly professional, commanding, and analytical tone. Do not use generic introductions.
    """

    final_summary = ""
    max_retries = 3

    # Retry Loop for API Stability
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt) # type: ignore
            final_summary = response.text.strip()
            break  # Success, exit the retry loop
        except Exception as e:
            logger.warning(f"Executive Summary API attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(2)  # Wait 2 seconds before trying again
            else:
                logger.error(f"Conversational Agent Error: {str(e)}")
                final_summary = f"Resolution executed successfully. {missing_item} rerouted using {assigned_fleet}. (Full AI summary failed to generate after {max_retries} attempts)."

    try:
        # Send Email via SMTP
        sender_email = os.environ.get("SENDER_EMAIL")
        sender_password = os.environ.get("SENDER_PASSWORD")
        
        if sender_email and sender_password and target_email:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = target_email
            msg['Subject'] = f"RESOLVED: SC001 Exception Report ({missing_item})"
            
            body = f"Stakeholder,\n\nThe autonomous swarm has resolved the recent supply chain exception. Please review the executive post-mortem below:\n\n{final_summary}"
            msg.attach(MIMEText(body, 'plain'))

            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, target_email, msg.as_string())
            server.quit()
            logger.info(f"Executive summary successfully emailed to {target_email}")
        else:
            logger.warning("Email not sent: Missing SMTP credentials or target email.")

        return final_summary

    except Exception as e:
        logger.error(f"SMTP Email Error: {str(e)}")
        return final_summary

def run_chat_interaction(email_body: str, history: list, user_message: str):
    """Acts as a live chatbot to debate solutions before execution, with automatic retries."""
    api_key = get_api_key()
    if not api_key:
        return "System Error: API Key missing."

    client = g_genai.Client(api_key=api_key) # type: ignore

    prompt = f"""
    You are an expert Supply Chain AI. The human overseer has paused an automated resolution to intervene manually.
    
    Here is the intercepted exception email:
    ---
    {email_body}
    ---
    
    INSTRUCTIONS:
    1. If the human says they are taking manual control, briefly state the exact issue found in the email, then offer 2-3 logical options.
    2. Be concise, professional, and do not use generic AI greetings. 
    3. Debate the options with the user if they ask questions.
    4. CRITICAL: Once the user explicitly confirms a plan or gives a final command to proceed, you MUST append the exact text "[EXECUTE]" at the very end of your response. This signals the system software to take over.
    """

    if history:
        prompt += "\n\nCHAT HISTORY:\n"
        for msg in history:
            role = "Human" if msg.get("sender") == "human" else "AI"
            prompt += f"{role}: {msg.get('text')}\n"

    prompt += f"\nHuman: {user_message}\nAI:"

    max_retries = 3

    # Retry Loop for API Stability
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt) # type: ignore
            return response.text.strip()
        except Exception as e:
            logger.warning(f"Chat API attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(2)  # Wait 2 seconds before trying again
            else:
                logger.error(f"Chat Agent Error after {max_retries} attempts: {str(e)}")
                return f"Error connecting to neural link after {max_retries} automatic retries. The network is currently congested. Please try again."