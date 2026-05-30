import os
import tempfile
from langchain_community.document_loaders import PyPDFLoader, CSVLoader, UnstructuredExcelLoader

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
            loader = PyPDFLoader(tmp_path)
        elif ext == "csv":
            loader = CSVLoader(tmp_path)
        elif ext in ["xlsx", "xls"]:
            loader = UnstructuredExcelLoader(tmp_path, mode="elements")

        docs = loader.load()
        doc_type = detect_doc_type(docs)
        return docs, doc_type

    finally:
        os.unlink(tmp_path)  


def detect_doc_type(docs):
    text = " ".join([d.page_content for d in docs]).lower()

    scores = {
        "balance_sheet":    sum(text.count(k) for k in ["total assets", "liabilities", "equity", "retained earnings", "shareholder"]),
        "profit_loss":      sum(text.count(k) for k in ["revenue", "gross profit", "ebitda", "net income", "operating expense"]),
        "cash_flow":        sum(text.count(k) for k in ["cash flow", "operating activities", "investing activities", "financing activities", "free cash"]),
    }

    detected = max(scores, key=scores.get)
    return detected if scores[detected] > 0 else "unknown"