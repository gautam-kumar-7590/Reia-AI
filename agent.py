# agent.py
import os
from dotenv import load_dotenv
from langchain_core.messages import HumanMessage, AIMessage
from llm_handler import get_llm
from Memory import recall_company, save_company
from rag import analyze_financials
from web_search import search_financial_news, search_company_data, search_industry_benchmark
from prompts import build_full_prompt, get_command_prompt
from langgraph.prebuilt import create_react_agent

load_dotenv()

# ─── REAL-TIME DATA KEYWORDS ──────────────────────────────────────────────────
REALTIME_KEYWORDS = [
    "exchange rate", "usd", "inr", "eur", "gbp", "price today", "current price",
    "stock price", "rupee", "dollar", "euro", "pound", "forex", "currency",
    "interest rate", "inflation", "gold price", "oil price", "crypto", "bitcoin",
    "market", "news", "latest", "current", "today", "right now", "live rate",
    "how much is", "what is the price", "what is the rate",
]

def needs_realtime_search(query: str) -> bool:
    q = query.lower()
    return any(kw in q for kw in REALTIME_KEYWORDS)

def get_realtime_context(query: str) -> str:
    """Call Tavily directly and return results as context string."""
    try:
        from web_search import get_search_tool
        results = get_search_tool().invoke(query)
        if not results:
            return ""
        lines = ["[LIVE WEB SEARCH RESULTS — use these numbers in your response]"]
        for r in results[:4]:
            lines.append(f"Source: {r.get('url', '')}")
            lines.append(r.get('content', ''))
            lines.append("---")
        return "\n".join(lines)
    except Exception as e:
        return f"[Search unavailable: {e}]"



# ─── TOOLS ────────────────────────────────────────────────────────────────────
def get_tools():
    return [
        recall_company,
        save_company,
        search_financial_news,
        search_company_data,
        search_industry_benchmark,
    ]


# ─── AGENT BUILDER ────────────────────────────────────────────────────────────
def build_agent(personality: str = "Professional", tone: str = "Analyst"):
    llm   = get_llm(streaming=False)
    tools = get_tools()

    system_prompt = build_full_prompt(personality=personality, tone=tone, user_input="")

    system_prompt += """

TOOL RULES — follow exactly:

SEARCH TOOLS (search_financial_news, search_company_data, search_industry_benchmark):
- ALWAYS call a search tool when the user asks about: current prices, exchange rates,
  stock prices, interest rates, inflation, news, benchmarks, or any real-world data.
- NEVER fabricate exchange rates, prices, or financial news. If you don't know the 
  current value, SEARCH — do not guess.
- Examples that REQUIRE a search tool:
  "what is USD to INR today" → search_financial_news("USD INR exchange rate today")
  "rupee vs euro current rate" → search_financial_news("INR EUR exchange rate 2025")
  "what is the current gold price" → search_financial_news("gold price today 2025")
  "news about Zerodha" → search_company_data("Zerodha")

FILE ANALYSIS:
- When file content is provided in the user message, analyze it DIRECTLY — no tool needed.
- Use exact numbers from the data. Never say you need a file when one is already provided.

MEMORY TOOLS:
- Call recall_company ONLY if user explicitly says "load [company name]" or "recall [company]"
- Call save_company ONLY if user says "save as [name]" or "remember this company"

NEVER:
- Call any tool with None, empty string, or missing required values
- Make up financial data or pretend to search without calling the tool
- Guess or assume a company name
"""

    agent = create_react_agent(model=llm, tools=tools, prompt=system_prompt)
    return agent


# ─── CHAT RUNNER ──────────────────────────────────────────────────────────────
def run_agent(
    query: str,
    chat_history: list,
    personality: str = "Professional",
    tone: str = "Analyst",
    docs=None,
    doc_type: str = None,
) -> str:

    command_prompt = get_command_prompt(query)
    full_query = f"{command_prompt}\n\n{query}".strip() if command_prompt else query

    # ── Inject file content into the message ──────────────────────────────────
    if docs:
        # 50K chars — handles large datasets without truncating key rows
        file_text = "\n\n".join([d.page_content for d in docs])
        file_text = file_text[:20000]

        full_query = f"""The user has uploaded a file for analysis.

DOCUMENT TYPE: {doc_type or 'unknown'}

FILE CONTENT (ALL ROWS):
{file_text}

---
USER REQUEST: {full_query}

Analyze the file content above. Use exact numbers from the data. 
Do not say you need a file — it is already provided above."""

    # ── Direct Tavily injection for real-time queries ────────────────────────
    if needs_realtime_search(query):
        realtime_ctx = get_realtime_context(query)
        if realtime_ctx:
            full_query = f"""{realtime_ctx}

---
USER REQUEST: {full_query}

Use the live search results above to answer. Give the exact current numbers. Do NOT say you lack real-time access."""

    agent    = build_agent(personality=personality, tone=tone)
    messages = chat_history + [HumanMessage(content=full_query)]

    try:
        result = agent.invoke({"messages": messages})
        output_messages = result.get("messages", [])
        for msg in reversed(output_messages):
            if isinstance(msg, AIMessage) and msg.content:
                return msg.content
        return "Reia could not generate a response. Please try again."

    except Exception as e:
        error_msg = str(e)
        try:
            llm      = get_llm(streaming=False)
            fallback = llm.invoke([HumanMessage(content=full_query)])
            return fallback.content
        except Exception as fallback_err:
            return f"Error: {error_msg} | Fallback error: {str(fallback_err)}"


if __name__ == "__main__":
    history = []
    print("Reia Agent ready. Type 'exit' to quit.\n")
    while True:
        user_input = input("You: ")
        if user_input.lower() == "exit":
            break
        response = run_agent(user_input, history)
        print(f"\nReia: {response}\n")
        history.append(HumanMessage(content=user_input))
        history.append(AIMessage(content=response))
