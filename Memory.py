# Memory.py
import sqlite3
import json
import os
from datetime import datetime
from langchain_core.tools import tool

DB_PATH = "reia_memory.db"

# ─── INIT DB ──────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS companies (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT UNIQUE NOT NULL,
            doc_type    TEXT,
            summary     TEXT,
            raw_data    TEXT,
            saved_at    TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()


# ─── CORE FUNCTIONS ───────────────────────────────────────────────────────────

def save_company_data(name: str, summary: str, doc_type: str = "unknown", raw_data: dict = {}):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO companies (name, doc_type, summary, raw_data, saved_at)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(name) DO UPDATE SET
            doc_type = excluded.doc_type,
            summary  = excluded.summary,
            raw_data = excluded.raw_data,
            saved_at = excluded.saved_at
    """, (
        name.lower().strip(),
        doc_type,
        summary,
        json.dumps(raw_data),
        datetime.now().strftime("%Y-%m-%d %H:%M")
    ))
    conn.commit()
    conn.close()
    return f"✅ Saved: {name}"


def load_company_data(name: str):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name, doc_type, summary, raw_data, saved_at FROM companies WHERE name = ?",
        (name.lower().strip(),)
    )
    row = cursor.fetchone()
    conn.close()

    if not row:
        return None

    return {
        "name":     row[0],
        "doc_type": row[1],
        "summary":  row[2],
        "raw_data": json.loads(row[3]),
        "saved_at": row[4],
    }


def list_all_companies():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT name, doc_type, saved_at FROM companies ORDER BY saved_at DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows


def delete_company(name: str):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM companies WHERE name = ?", (name.lower().strip(),))
    conn.commit()
    conn.close()
    return f"🗑️ Deleted: {name}"


# ─── LANGCHAIN TOOLS (plugs into agent.py) ────────────────────────────────────

@tool
def recall_company(company_name: str) -> str:
    """Recall previously saved financial data for a company from memory."""
    data = load_company_data(company_name)
    if not data:
        return f"No saved data found for '{company_name}'."
    return (
        f"Company: {data['name'].title()}\n"
        f"Type: {data['doc_type']}\n"
        f"Saved: {data['saved_at']}\n\n"
        f"Summary:\n{data['summary']}"
    )


@tool
def save_company(company_name: str, summary: str) -> str:
    """Save financial analysis summary for a company to memory."""
    return save_company_data(name=company_name, summary=summary)