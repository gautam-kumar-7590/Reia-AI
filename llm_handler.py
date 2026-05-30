import os
from dotenv import load_dotenv

load_dotenv()

SYSTEM_PROMPT = """You are Reia, a sharp AI financial analyst. 
Analyze financial documents with precision. Give brutal honest verdicts. 
Always back insights with numbers from the data."""

def get_llm(streaming=False):
    from langchain_groq import ChatGroq
    return ChatGroq(
        api_key=os.getenv("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"),
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        streaming=streaming,
    )