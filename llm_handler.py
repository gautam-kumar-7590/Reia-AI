# llm_handler.py
import os
from dotenv import load_dotenv

load_dotenv()


def get_llm(streaming=False):
    """
    Priority order:
    1. Ollama via ngrok (local Qwen2.5 14B) — if OLLAMA_BASE_URL is set
    2. Groq fallback — only if Ollama fails or URL not set
    """

    ollama_url = os.getenv("OLLAMA_BASE_URL", "").strip()

    if ollama_url:
        try:
            from langchain_ollama import ChatOllama
            llm = ChatOllama(
                base_url=ollama_url,
                model=os.getenv("OLLAMA_MODEL", "qwen2.5:14b"),
                temperature=0.3,
                streaming=streaming,
                timeout=180,
            )
            # Verify connection with a lightweight check
            import requests
            resp = requests.get(f"{ollama_url}/api/tags", timeout=10)
            if resp.status_code == 200:
                print(f"[LLM] Using Ollama @ {ollama_url}")
                return llm
            else:
                raise Exception(f"Ollama returned {resp.status_code}")
        except Exception as e:
            print(f"[LLM] Ollama failed ({e}) — falling back to Groq")

    # Groq fallback
    groq_key = os.getenv("GROQ_API_KEY", "").strip()
    if not groq_key:
        raise ValueError(
            "No LLM available. Set OLLAMA_BASE_URL for local model, "
            "or GROQ_API_KEY for cloud fallback."
        )

    from langchain_groq import ChatGroq
    print("[LLM] Using Groq fallback")
    return ChatGroq(
        api_key=groq_key,
        model="llama-3.3-70b-versatile",
        temperature=0.3,
        streaming=streaming,
    )
