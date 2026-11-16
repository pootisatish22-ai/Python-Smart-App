"""Quick connection test for whichever AI provider is set in .env.   Run:  python test_llm.py"""
import os
import time

from dotenv import load_dotenv
load_dotenv()
from core import get_llm, llm_ready

print("Provider:", os.getenv("LLM_PROVIDER", "gemini"))
problem = llm_ready()
if problem:
    raise SystemExit("PROBLEM: " + problem)
t = time.time()
try:
    r = get_llm().invoke("Reply with the single word: OK")
    print("Reply:", r.content, "| took", round(time.time() - t, 1), "seconds")
    print("SUCCESS - you can run: streamlit run app.py")
except Exception as e:
    print("FAILED after", round(time.time() - t, 1), "seconds:")
    print(e)
