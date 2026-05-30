# app.py
import streamlit as st
from langchain_core.messages import HumanMessage, AIMessage
from file_reader import read_file
from rag import load_docs_to_chroma
from agent import run_agent
from Memory import list_all_companies, delete_company
from ExcelEngine import generate_excel_report
import os

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Reia — AI Financial Analyst",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── THEME CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Geist:wght@300;400;500;600&family=Geist+Mono:wght@400;500&display=swap');

    /* ── Base ── */
    html, body, .stApp {
        background-color: #1a1f2e !important;
        color: #e8e9ed !important;
        font-family: 'Geist', -apple-system, sans-serif !important;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background-color: #111520 !important;
        border-right: 0.5px solid rgba(255,255,255,0.06) !important;
    }
    [data-testid="stSidebar"] * { font-family: 'Geist', sans-serif !important; }

    /* Sidebar logo area */
    .reia-logo {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 4px 0 20px;
    }
    .reia-logo-icon {
        width: 32px; height: 32px;
        background: linear-gradient(135deg, #F0A500, #e8930a);
        border-radius: 8px;
        display: flex; align-items: center; justify-content: center;
        font-size: 16px; font-weight: 600; color: #111520;
    }
    .reia-logo-text { font-size: 17px; font-weight: 600; color: #e8e9ed; letter-spacing: -0.3px; }

    /* Sidebar section labels */
    .sidebar-label {
        font-size: 11px;
        font-weight: 500;
        color: rgba(232,233,237,0.4);
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin: 16px 0 8px;
    }

    /* ── Chat messages ── */
    [data-testid="stChatMessage"] {
        background: transparent !important;
        border: none !important;
        padding: 4px 0 !important;
    }

    /* User messages */
    [data-testid="stChatMessage"]:has([data-testid="stChatMessageContent"]):has(.user-msg) {
        justify-content: flex-end;
    }

    .stChatMessage [data-testid="stMarkdownContainer"] p {
        font-size: 15px !important;
        line-height: 1.65 !important;
        color: #e8e9ed !important;
    }

    /* User bubble */
    [data-testid="stChatMessage"][aria-label="user"] [data-testid="stMarkdownContainer"] {
        background: #2a3050 !important;
        border-radius: 18px 18px 4px 18px !important;
        padding: 10px 16px !important;
        max-width: 80% !important;
        margin-left: auto !important;
        border: 0.5px solid rgba(240,165,0,0.15) !important;
    }

    /* Assistant message — no bubble, full width */
    [data-testid="stChatMessage"][aria-label="assistant"] [data-testid="stMarkdownContainer"] {
        padding: 2px 4px !important;
    }

    /* ── Chat input bar ── */
    [data-testid="stBottom"] {
        background: transparent !important;
        padding-bottom: 12px !important;
    }
    [data-testid="stChatInput"] {
        background: #222841 !important;
        border: 0.5px solid rgba(255,255,255,0.1) !important;
        border-radius: 16px !important;
        padding: 4px 8px !important;
        box-shadow: 0 0 0 0px transparent !important;
        transition: border-color 0.15s !important;
    }
    [data-testid="stChatInput"]:focus-within {
        border-color: rgba(240,165,0,0.4) !important;
    }
    [data-testid="stChatInput"] textarea {
        background: transparent !important;
        color: #e8e9ed !important;
        font-family: 'Geist', sans-serif !important;
        font-size: 15px !important;
        caret-color: #F0A500 !important;
    }
    [data-testid="stChatInput"] textarea::placeholder { color: rgba(232,233,237,0.35) !important; }

    /* Send button */
    [data-testid="stChatInput"] button {
        background: #F0A500 !important;
        border-radius: 10px !important;
        border: none !important;
        color: #111520 !important;
    }
    [data-testid="stChatInput"] button:hover { background: #00D4FF !important; }

    /* ── Mode selector (Claude-style, above input) ── */
    .mode-bar {
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 6px 2px 4px;
    }
    .mode-label {
        font-size: 12px;
        color: rgba(232,233,237,0.4);
        font-weight: 400;
    }

    /* Streamlit selectbox styled as pill */
    [data-testid="stSelectbox"] > div > div {
        background: #222841 !important;
        border: 0.5px solid rgba(255,255,255,0.1) !important;
        border-radius: 20px !important;
        color: #e8e9ed !important;
        font-size: 13px !important;
        font-family: 'Geist', sans-serif !important;
        min-height: 32px !important;
        padding: 0 12px !important;
    }
    [data-testid="stSelectbox"] > div > div:hover {
        border-color: rgba(240,165,0,0.3) !important;
    }
    [data-testid="stSelectbox"] svg { color: rgba(232,233,237,0.4) !important; }

    /* ── Buttons ── */
    .stButton > button {
        background: transparent !important;
        color: rgba(232,233,237,0.7) !important;
        border: 0.5px solid rgba(255,255,255,0.1) !important;
        border-radius: 8px !important;
        font-size: 12.5px !important;
        font-family: 'Geist', sans-serif !important;
        font-weight: 400 !important;
        padding: 6px 12px !important;
        transition: all 0.15s !important;
    }
    .stButton > button:hover {
        border-color: rgba(240,165,0,0.35) !important;
        color: #F0A500 !important;
        background: rgba(240,165,0,0.06) !important;
    }

    /* Primary CTA button */
    .stButton.primary > button {
        background: #F0A500 !important;
        color: #111520 !important;
        border: none !important;
        font-weight: 500 !important;
    }
    .stButton.primary > button:hover { background: #00D4FF !important; }

    /* ── File uploader ── */
    [data-testid="stFileUploader"] {
        background: #1e2335 !important;
        border: 1px dashed rgba(240,165,0,0.25) !important;
        border-radius: 12px !important;
    }
    [data-testid="stFileUploader"] label { color: rgba(232,233,237,0.5) !important; font-size: 13px !important; }

    /* ── Text input ── */
    [data-testid="stTextInput"] input {
        background: #1e2335 !important;
        color: #e8e9ed !important;
        border: 0.5px solid rgba(255,255,255,0.1) !important;
        border-radius: 8px !important;
        font-family: 'Geist', sans-serif !important;
        font-size: 14px !important;
    }
    [data-testid="stTextInput"] input:focus { border-color: rgba(240,165,0,0.4) !important; }

    /* ── Metrics ── */
    [data-testid="stMetric"] {
        background: #1e2335 !important;
        border: 0.5px solid rgba(255,255,255,0.07) !important;
        border-radius: 10px !important;
        padding: 12px 16px !important;
    }
    [data-testid="stMetric"] label { font-size: 11px !important; color: rgba(232,233,237,0.4) !important; text-transform: uppercase; letter-spacing: 0.06em; }
    [data-testid="stMetric"] [data-testid="stMetricValue"] { font-size: 16px !important; font-weight: 500 !important; color: #e8e9ed !important; }

    /* ── Divider ── */
    hr { border-color: rgba(255,255,255,0.06) !important; margin: 12px 0 !important; }

    /* ── Status badge ── */
    .status-badge {
        display: inline-flex; align-items: center; gap: 5px;
        background: rgba(0,212,255,0.08);
        border: 0.5px solid rgba(0,212,255,0.2);
        border-radius: 20px;
        padding: 3px 10px;
        font-size: 12px; font-weight: 500; color: #00D4FF;
    }
    .status-dot { width: 6px; height: 6px; border-radius: 50%; background: #00D4FF; }

    /* ── Quick command pills ── */
    .quick-pill {
        display: inline-flex; align-items: center; gap: 4px;
        font-size: 12px !important;
    }

    /* ── Scrollbar ── */
    ::-webkit-scrollbar { width: 4px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); border-radius: 4px; }

    /* ── Headings ── */
    h1, h2, h3 { color: #e8e9ed !important; font-family: 'Geist', sans-serif !important; font-weight: 500 !important; letter-spacing: -0.3px !important; }

    /* ── Success / error alerts ── */
    [data-testid="stAlert"] { border-radius: 10px !important; font-size: 13px !important; }

    /* ── Hide streamlit branding ── */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    [data-testid="stHeader"] { background: transparent !important; }

    /* ── Spinner ── */
    [data-testid="stSpinner"] { color: #F0A500 !important; }

    /* ── Saved company pills in sidebar ── */
    .company-pill {
        display: flex; align-items: center; justify-content: space-between;
        background: #1e2335;
        border: 0.5px solid rgba(255,255,255,0.07);
        border-radius: 8px;
        padding: 8px 10px;
        margin-bottom: 5px;
    }
    .company-pill-name { font-size: 13px; font-weight: 500; color: #e8e9ed; }
    .company-pill-type { font-size: 11px; color: rgba(232,233,237,0.4); margin-top: 1px; }

    /* ── Main header area ── */
    .main-header {
        display: flex; align-items: center; justify-content: space-between;
        padding: 8px 0 16px;
        border-bottom: 0.5px solid rgba(255,255,255,0.06);
        margin-bottom: 20px;
    }
    .main-title { font-size: 20px; font-weight: 500; color: #e8e9ed; letter-spacing: -0.4px; }
    .main-subtitle { font-size: 13px; color: rgba(232,233,237,0.4); margin-top: 2px; }

    /* Empty state */
    .empty-state {
        display: flex; flex-direction: column; align-items: center;
        justify-content: center; padding: 60px 20px; text-align: center;
    }
    .empty-icon { font-size: 40px; margin-bottom: 16px; opacity: 0.5; }
    .empty-title { font-size: 18px; font-weight: 500; color: rgba(232,233,237,0.7); margin-bottom: 8px; }
    .empty-sub { font-size: 14px; color: rgba(232,233,237,0.35); max-width: 340px; line-height: 1.6; }
</style>
""", unsafe_allow_html=True)


# ─── SESSION STATE ────────────────────────────────────────────────────────────
defaults = {
    "chat_history": [],
    "docs_loaded": False,
    "doc_type": "unknown",
    "company_name": "",
    "personality": "Professional",
    "tone": "Analyst",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─── PERSONALITY CONFIG ───────────────────────────────────────────────────────
MODES = {
    "Professional": {"icon": "📊", "desc": "Formal analyst mode", "color": "#00D4FF"},
    "Chika":        {"icon": "🎀", "desc": "Bubbly & cheerful",   "color": "#F0A500"},
    "Rei":          {"icon": "🔵", "desc": "Calm & precise",       "color": "#378ADD"},
    "Toga":         {"icon": "🔪", "desc": "Brutal & unfiltered",  "color": "#E24B4A"},
}


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    # Logo
    st.markdown("""
    <div class="reia-logo">
        <div class="reia-logo-icon">R</div>
        <div class="reia-logo-text">Reia CFO</div>
    </div>
    """, unsafe_allow_html=True)

    # Upload section
    st.markdown('<div class="sidebar-label">Document</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Upload",
        type=["xlsx", "xls", "pdf", "csv"],
        label_visibility="collapsed"
    )

    if uploaded_file:
        company_input = st.text_input("Company name", placeholder="e.g. Zerodha", label_visibility="collapsed")
        if st.button("Analyze →", use_container_width=True):
            with st.spinner("Reading..."):
                try:
                    docs, doc_type = read_file(uploaded_file)
                    load_docs_to_chroma(docs)
                    st.session_state.docs_loaded = True
                    st.session_state.doc_type = doc_type
                    st.session_state.company_name = company_input or uploaded_file.name
                    st.success(f"Loaded: {doc_type.replace('_', ' ').title()}")
                except Exception as e:
                    st.error(str(e))

    st.divider()

    # Tone
    st.markdown('<div class="sidebar-label">Response tone</div>', unsafe_allow_html=True)
    tone = st.selectbox(
        "Tone",
        ["Analyst", "Casual", "Brutal", "Report"],
        label_visibility="collapsed"
    )
    st.session_state.tone = tone

    st.divider()

    # Saved companies
    st.markdown('<div class="sidebar-label">Saved companies</div>', unsafe_allow_html=True)
    companies = list_all_companies()
    if companies:
        for name, doc_type, saved_at in companies:
            col1, col2 = st.columns([4, 1])
            col1.markdown(f"""
            <div style="padding: 2px 0;">
                <div style="font-size:13px; font-weight:500; color:#e8e9ed;">{name.title()}</div>
                <div style="font-size:11px; color:rgba(232,233,237,0.4);">{doc_type}</div>
            </div>
            """, unsafe_allow_html=True)
            if col2.button("✕", key=f"del_{name}"):
                delete_company(name)
                st.rerun()
    else:
        st.markdown('<div style="font-size:13px; color:rgba(232,233,237,0.3); padding: 4px 0;">No companies saved</div>', unsafe_allow_html=True)

    st.divider()

    # Clear chat
    if st.button("Clear chat", use_container_width=True):
        st.session_state.chat_history = []
        st.rerun()


# ─── MAIN AREA ────────────────────────────────────────────────────────────────

# Header
header_right = ""
if st.session_state.docs_loaded:
    header_right = f'<span class="status-badge"><span class="status-dot"></span>{st.session_state.doc_type.replace("_"," ").title()}</span>'

st.markdown(f"""
<div class="main-header">
    <div>
        <div class="main-title">Reia — AI Financial Analyst</div>
        <div class="main-subtitle">Your personal AI CFO. Upload a financial document to get started.</div>
    </div>
    <div>{header_right}</div>
</div>
""", unsafe_allow_html=True)

# Metrics row (shown only when doc loaded)
if st.session_state.docs_loaded:
    m1, m2, m3 = st.columns(3)
    m1.metric("Company", st.session_state.company_name or "—")
    m2.metric("Document", st.session_state.doc_type.replace("_", " ").title())
    m3.metric("Mode", f"{MODES[st.session_state.personality]['icon']} {st.session_state.personality}")
    st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

# Quick commands (only when doc loaded)
quick_command = None
if st.session_state.docs_loaded:
    st.markdown('<div style="font-size:12px; color:rgba(232,233,237,0.35); margin-bottom:6px;">Quick commands</div>', unsafe_allow_html=True)
    qc1, qc2, qc3, qc4, qc5 = st.columns(5)
    if qc1.button("🔥 Roast"):     quick_command = "roast this company"
    if qc2.button("🚩 Red Flags"): quick_command = "just red flags"
    if qc3.button("💸 Debt"):      quick_command = "focus on debt"
    if qc4.button("⚖️ Verdict"):   quick_command = "give me the verdict"
    if qc5.button("📊 Compare"):   quick_command = "compare companies"
    st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

# ─── CHAT HISTORY ─────────────────────────────────────────────────────────────
if not st.session_state.chat_history:
    st.markdown("""
    <div class="empty-state">
        <div class="empty-icon">📊</div>
        <div class="empty-title">Ask Reia anything</div>
        <div class="empty-sub">Upload a Balance Sheet, P&L, or Cash Flow statement from the sidebar to unlock full analysis. Or just ask a financial question.</div>
    </div>
    """, unsafe_allow_html=True)
else:
    for message in st.session_state.chat_history:
        if isinstance(message, HumanMessage):
            with st.chat_message("user"):
                st.markdown(message.content)
        elif isinstance(message, AIMessage):
            with st.chat_message("assistant", avatar=MODES[st.session_state.personality]["icon"]):
                st.markdown(message.content)


# ─── MODE SELECTOR + CHAT INPUT ───────────────────────────────────────────────
# Mode selector row — sits just above the chat input, mimics Claude's model picker
col_label, col_mode, col_spacer = st.columns([0.12, 0.28, 0.6])

with col_label:
    st.markdown('<div style="font-size:12px; color:rgba(232,233,237,0.35); padding-top:8px;">Mode</div>', unsafe_allow_html=True)

with col_mode:
    mode_options = list(MODES.keys())
    mode_labels  = [f"{MODES[m]['icon']}  {m}" for m in mode_options]
    current_idx  = mode_options.index(st.session_state.personality)
    selected_label = st.selectbox(
        "Mode",
        mode_labels,
        index=current_idx,
        label_visibility="collapsed",
    )
    selected_mode = mode_options[mode_labels.index(selected_label)]
    st.session_state.personality = selected_mode

# Chat input
user_input = st.chat_input(
    f"Message Reia ({selected_mode})…"
) or quick_command

# ─── RESPONSE ─────────────────────────────────────────────────────────────────
if user_input:
    with st.chat_message("user"):
        st.markdown(user_input)

    with st.chat_message("assistant", avatar=MODES[st.session_state.personality]["icon"]):
        with st.spinner("Analyzing…"):
            try:
                response = run_agent(
                    query=user_input,
                    chat_history=st.session_state.chat_history,
                    personality=st.session_state.personality,
                )
                st.markdown(response)

                # Excel download
                if any(w in user_input.lower() for w in ["analyze", "verdict", "report", "roast"]):
                    if st.button("Download Excel Report →"):
                        path = generate_excel_report(
                            company_name=st.session_state.company_name or "Company",
                            doc_type=st.session_state.doc_type,
                            ratios={},
                            positives=["See full chat analysis"],
                            negatives=["See full chat analysis"],
                            insights=[response[:200]],
                            actions=[],
                            score=0,
                            verdict=response[:100],
                        )
                        with open(path, "rb") as f:
                            st.download_button(
                                "📥 Download Excel",
                                f,
                                file_name=os.path.basename(path),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

            except Exception as e:
                st.error(str(e))
                response = f"Error: {str(e)}"

    st.session_state.chat_history.append(HumanMessage(content=user_input))
    st.session_state.chat_history.append(AIMessage(content=response))