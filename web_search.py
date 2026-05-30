# web_search.py
import os
from langchain_community.tools.tavily_search import TavilySearchResults
from langchain_core.tools import tool
from dotenv import load_dotenv

load_dotenv()

# ─── TAVILY SEARCH TOOL ───────────────────────────────────────────────────────
# CORRECT
def get_search_tool():
    api_key = os.getenv("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")
    if not api_key:
        raise ValueError("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")
    return TavilySearchResults(
        max_results=5,
        search_depth="advanced",
        include_answer=True,
        include_raw_content=False,
    )
# ─── FINANCIAL SPECIFIC SEARCH ────────────────────────────────────────────────
@tool
def search_financial_news(query: str) -> str:
    """Search the web for latest financial news, market data, or company information."""
    try:
        search = get_search_tool()
        results = search.invoke(query)

        if not results:
            return "No results found."

        output = []
        for r in results:
            output.append(
                f"📰 {r.get('title', 'No Title')}\n"
                f"🔗 Source: {r.get('url', '')}\n"
                f"📝 {r.get('content', '')}\n"
            )

        return "\n---\n".join(output)

    except Exception as e:
        return f"Search error: {str(e)}"


# ─── COMPANY SPECIFIC SEARCH ──────────────────────────────────────────────────
@tool
def search_company_data(company_name: str) -> str:
    """Search for latest data, news, and financials about a specific company."""
    try:
        search = get_search_tool()
        query = f"{company_name} financials revenue profit 2024 2025"
        results = search.invoke(query)

        if not results:
            return f"No data found for {company_name}."

        output = []
        for r in results:
            output.append(
                f"📰 {r.get('title', 'No Title')}\n"
                f"🔗 Source: {r.get('url', '')}\n"
                f"📝 {r.get('content', '')}\n"
            )

        return "\n---\n".join(output)

    except Exception as e:
        return f"Search error: {str(e)}"


# ─── BENCHMARK SEARCH ─────────────────────────────────────────────────────────
@tool
def search_industry_benchmark(industry: str, metric: str) -> str:
    """Search for industry benchmark data for a specific metric like ROE, D/E ratio etc."""
    try:
        search = get_search_tool()
        query = f"{industry} industry average {metric} benchmark 2024 2025"
        results = search.invoke(query)

        if not results:
            return f"No benchmark data found for {industry} - {metric}."

        output = []
        for r in results:
            output.append(
                f"📰 {r.get('title', 'No Title')}\n"
                f"🔗 Source: {r.get('url', '')}\n"
                f"📝 {r.get('content', '')}\n"
            )

        return "\n---\n".join(output)

    except Exception as e:
        return f"Search error: {str(e)}"