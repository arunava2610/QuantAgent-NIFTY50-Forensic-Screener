import os
import imaplib
import email
import email.utils
from fastapi import APIRouter
from pydantic import BaseModel
from .conversational_agent import generate_executive_summary, run_chat_interaction

import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import logger

from .orchestrator_agent import run_orchestrator
from .node_utility import run_node_allocation
from .fleet_utility import run_fleet_assignment

router = APIRouter(prefix="/api/supply-chain", tags=["Supply Chain Swarm"])


class ExecuteRequest(BaseModel):
    email_body: str
    sender_email: str
    human_command: str = "" 

def scan_inbox_for_sc001():
    user = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_PASSWORD") 
    
    if not user or not password:
        return None, None, "IMAP credentials missing."

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(user, password)
        mail.select("inbox")

        status, messages = mail.search(None, '(UNSEEN SUBJECT "SC001")')
        mail_ids = messages[0].split()

        if not mail_ids:
            mail.logout()
            return None, None, "No unread SC001 exception flags found."

        latest_email_id = mail_ids[-1]
        status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
        
        body = ""
        sender_address = ""
        
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                
                # Extract the exact sender email
                raw_sender = msg.get("From")
                sender_address = email.utils.parseaddr(raw_sender)[1]

                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode()
                            break
                else:
                    body = msg.get_payload(decode=True).decode()

        mail.store(latest_email_id, '+FLAGS', '\\Seen') 
        mail.logout()
        
        return body.strip(), sender_address, "Success"
    except Exception as e:
        return None, None, f"IMAP Error: {str(e)}"

@router.get("/poll")
def poll_inbox():
    email_body, sender_email, status_msg = scan_inbox_for_sc001()
    if not email_body:
        return {"status": "idle", "message": status_msg}
        
    logger.info(f"[ACTION] 📥 Match found from {sender_email}. Marked as Read.")
    return {"status": "exception_found", "email_body": email_body, "sender_email": sender_email}

class ChatRequest(BaseModel):
    email_body: str
    history: list
    user_message: str

@router.post("/chat")
def chat_intervention(payload: ChatRequest):
    """Handles the live human-in-the-loop chat window."""
    reply = run_chat_interaction(payload.email_body, payload.history, payload.user_message)
    return {"reply": reply}

@router.post("/execute")
def execute_swarm(payload: ExecuteRequest):
    logs = []
    def add_log(l_type, msg):
        logs.append({"type": l_type, "message": msg})
        logger.info(f"[{l_type.upper()}] {msg}")

    try:
        add_log("analysis", "🧠 Orchestrator Agent: Parsing parameters into structured JSON...")
        orchestrator_data = run_orchestrator(payload.email_body, payload.human_command)
        missing_item = orchestrator_data.get("missing_item", "Unknown Item")
        add_log("analysis", f"🧠 Orchestrator: Target missing item: {missing_item}. Waking up Node Allocation Utility...")

        add_log("action", "🛠️ Tool Call: Fetching Logistics_ERP_Simulation.xlsx from Google Drive.")
        allocated_node, reserved_qty = run_node_allocation(missing_item)
        
        if not allocated_node:
            add_log("system", f"❌ Node Allocation Failed: Could not find '{missing_item}'.")
            return {"logs": logs, "success": False}
            
        add_log("analysis", f"🧠 Node Allocation Agent: Stock located at {allocated_node}. Reserving {reserved_qty} units.")

        add_log("action", f"🛠️ Fleet Agent: Assigning asset and modifying in-memory Excel matrix.")
        assigned_fleet, b64_file = run_fleet_assignment(allocated_node)
        add_log("success", "🟩 [EXCEL ENGINE] Changes highlighted in GREEN. Encoding to Base64 for secure browser transfer.")

        add_log("analysis", f"🧠 Conversational Agent: Drafting executive summary and emailing {payload.sender_email}...")
        
        # Capture the generated text
        final_summary_text = generate_executive_summary(missing_item, allocated_node, assigned_fleet, payload.sender_email)
        
        # Inject the full text into the React UI logs
        add_log("success", f"✅ Execution Complete & Emailed. Executive Summary:\n\n{final_summary_text}")
        
        return {
            "logs": logs, 
            "success": True, 
            "file_b64": b64_file 
        }

    except Exception as e:
        logger.error(f"🚨 Swarm Execution Error: {str(e)}")
        add_log("system", f"❌ Pipeline Failed: {str(e)}")
        return {"logs": logs, "success": False}