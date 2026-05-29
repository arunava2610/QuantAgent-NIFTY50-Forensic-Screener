import os
from dotenv import load_dotenv, find_dotenv

# 🔎 1. Find the exact absolute path of the file Python is discovering
exact_env_path = find_dotenv()

print("========== 🕵️ FORENSIC ENV PORTAL ==========")
print(f"📍 EXACT FILE PATH BEING READ:\n   {exact_env_path}\n")

# 2. Force load it completely
load_dotenv(exact_env_path, override=True)

key = os.environ.get("GEMINI_API_KEY", "")
print(f"📏 Key Character Length: {len(key)}")
print(f"🔑 Key Text Starts With: {key[:7] if key else 'NOTHING_FOUND'}")
print("============================================")