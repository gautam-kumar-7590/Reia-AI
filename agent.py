# agent.py
import os
from dotenv import load_dotenv
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage
from llm_handler import get_llm
from Memory import recall_company, save_company
from rag import analyze_financials
from web_search import search_financial_news, search_company_data, search_industry_benchmark
from ExcelEngine import create_excel_report
from prompts import build_full_prompt, get_command_prompt
from langgraph.prebuilt import create_react_agent

load_dotenv()


# ─── TOOLS ────────────────────────────────────────────────────────────────────
def get_tools():
    return [
        analyze_financials,
        recall_company,
        save_company,
        search_financial_news,
        search_company_data,
        search_industry_benchmark,
        create_excel_report,
    ]


# ─── AGENT BUILDER ────────────────────────────────────────────────────────────
def build_agent(personality: str = "Professional", tone: str = "Analyst"):
    llm = get_llm(streaming=True)
    tools = get_tools()

    system_prompt = build_full_prompt(
        personality=personality,
        tone=tone,
        user_input=""
    )

    system_prompt += """

IMPORTANT TOOL RULES:
- For greetings, casual chat, or general questions → respond directly, NO tools
- Only call recall_company if user explicitly mentions a specific company name
- Only call save_company if user says "save" or "remember this company"
- Only call search tools if user explicitly asks for news, benchmarks, or industry data
- Only call analyze_financials if a financial file has been uploaded by the user
- Only call create_excel_report if user explicitly asks for an Excel report
- NEVER call any tool with None, empty string, or missing values
- NEVER guess or assume a company name — ask the user if unsure
"""

    agent = create_react_agent(
        model=llm,
        tools=tools,
        prompt=system_prompt,
    )

    return agent


# ─── CHAT RUNNER ──────────────────────────────────────────────────────────────
def run_agent(
    query: str,
    chat_history: list,
    personality: str = "Professional",
    tone: str = "Analyst"
) -> str:

    command_prompt = get_command_prompt(query)
    full_query = f"{command_prompt}\n\n{query}".strip() if command_prompt else query

    agent = build_agent(personality=personality, tone=tone)

    messages = chat_history + [HumanMessage(content=full_query)]

    try:
        result = agent.invoke({"messages": messages})

        output_messages = result.get("messages", [])
        for msg in reversed(output_messages):
            if isinstance(msg, AIMessage):
                return msg.content

        return "Reia could not generate a response."

    except Exception as e:
        error_msg = str(e)
        if "tool call validation" in error_msg or "not in request.tools" in error_msg:
            # Fallback: answer directly without tools
            llm = get_llm(streaming=False)
            fallback = llm.invoke([HumanMessage(content=full_query)])
            return fallback.content
        return f"Error: {error_msg}"


# ─── QUICK TEST ───────────────────────────────────────────────────────────────
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
