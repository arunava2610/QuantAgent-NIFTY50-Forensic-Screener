import json
import sys
import os
from google import genai as g_genai

# Import config from parent directory
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import logger, get_api_key

def run_orchestrator(email_body: str, human_command: str = ""):
    """Parses disruption reports and overrides into structured agentic commands."""
    api_key = get_api_key()
    if not api_key:
        raise Exception("Gemini API key missing or invalid in Google Drive.")

    client = g_genai.Client(api_key=api_key) # type: ignore

    prompt = f"""
    You are the Master Orchestrator for an autonomous supply chain swarm.
    Analyze this incident report:
    ---
    {email_body}
    ---
    """

    if human_command:
        prompt += f"\nCRITICAL HUMAN OVERRIDE: {human_command}\nYou MUST incorporate this human directive into your target resolution.\n"

    prompt += """
    Output STRICTLY valid JSON ONLY. No markdown, no backticks, no explanations. Use exactly these keys:
    {
      "missing_item": "Extract the specific item name or SKU needed (e.g., 'Enterprise Server Racks')",
      "urgency": "High/Medium/Low",
      "failed_asset": "Extract the truck ID, node, or asset that failed"
    }
    """

    response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt) # type: ignore
    cleaned_json = response.text.replace("```json", "").replace("```", "").strip()

    try:
        data = json.loads(cleaned_json)
        return data
    except Exception as e:
        logger.error(f"Orchestrator failed to parse JSON: {cleaned_json}")
        raise Exception("Orchestrator returned invalid data structure.")