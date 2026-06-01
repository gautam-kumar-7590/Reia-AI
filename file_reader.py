# file_reader.py
import os
import tempfile
import pandas as pd
import pdfplumber
from langchain_core.documents import Document

SUPPORTED = ["pdf", "csv", "xlsx", "xls"]


def read_file(uploaded_file):
    ext = uploaded_file.name.split(".")[-1].lower()

    if ext not in SUPPORTED:
        raise ValueError(f"Unsupported file type: .{ext}")

    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    try:
        if ext == "pdf":
            docs = load_pdf(tmp_path)
        elif ext == "csv":
            docs = load_csv(tmp_path)
        elif ext in ["xlsx", "xls"]:
            docs = load_excel(tmp_path)

        doc_type = detect_doc_type(docs)
        return docs, doc_type

    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def load_csv(path):
    df = None
    for enc in ["utf-8", "latin-1", "cp1252", "utf-8-sig"]:
        try:
            df = pd.read_csv(path, encoding=enc)
            break
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue

    if df is None:
        df = pd.read_csv(path, encoding="utf-8", errors="replace")

    total_rows = len(df)
    total_cols = len(df.columns)

    # For agent context: send summary + sample rows only (not all 51K rows)
    # Full DataFrame is read separately in app.py for ExcelEngine
    sample_size = min(200, total_rows)
    sample_df   = df.head(sample_size)

    summary = f"""Dataset Summary:
- Total Rows: {total_rows:,}
- Total Columns: {total_cols}
- Columns: {', '.join(df.columns.tolist())}
- Numeric columns: {', '.join(df.select_dtypes(include='number').columns.tolist())}

Column Stats:
{df.describe().to_string()}

Sample Data (first {sample_size} of {total_rows:,} rows):
{sample_df.to_string(index=False)}"""

    return [Document(
        page_content=summary,
        metadata={"source": path, "rows": total_rows, "cols": total_cols}
    )]


def load_excel(path):
    xl   = pd.ExcelFile(path)
    docs = []
    for sheet in xl.sheet_names:
        df         = xl.parse(sheet).fillna("")
        total_rows = len(df)
        sample_df  = df.head(200)
        text = f"""Sheet: {sheet}
Total Rows: {total_rows:,} | Columns: {', '.join(df.columns.astype(str).tolist())}

Sample Data (first {min(200, total_rows)} rows):
{sample_df.to_string(index=False)}"""
        docs.append(Document(
            page_content=text,
            metadata={"source": path, "sheet": sheet, "rows": total_rows}
        ))
    return docs


def load_pdf(path):
    docs = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if text.strip():
                docs.append(Document(
                    page_content=text,
                    metadata={"source": path, "page": i + 1}
                ))
    return docs if docs else [Document(
        page_content="Could not extract text from PDF.",
        metadata={"source": path}
    )]


def detect_doc_type(docs):
    text = " ".join([d.page_content for d in docs]).lower()

    scores = {
        "balance_sheet": sum(text.count(k) for k in ["total assets", "liabilities", "equity", "retained earnings", "shareholder"]),
        "profit_loss":   sum(text.count(k) for k in ["revenue", "gross profit", "ebitda", "net income", "operating expense"]),
        "cash_flow":     sum(text.count(k) for k in ["cash flow", "operating activities", "investing activities", "financing activities", "free cash"]),
        "sales_data":    sum(text.count(k) for k in ["sales", "quantity", "orders", "customer", "product", "category", "profit", "discount", "ship"]),
    }

    detected = max(scores, key=scores.get)
    return detected if scores[detected] > 0 else "unknown"
