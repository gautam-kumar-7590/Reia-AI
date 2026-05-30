# rag.py
import os
from langchain_community.vectorstores import Chroma
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_core.tools import tool
from langchain_core.output_parsers import StrOutputParser
from langchain_core.runnables import RunnablePassthrough
from llm_handler import get_llm

CHROMA_DIR = "chroma_db"

# ─── EMBEDDINGS ───────────────────────────────────────────────────────────────
def get_embeddings():
    from langchain_huggingface import HuggingFaceEmbeddings
    return HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2"
    )


# ─── LOAD DOCS INTO CHROMA ────────────────────────────────────────────────────
def load_docs_to_chroma(docs: list, collection_name: str = "reia_docs"):
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=150,
    )
    chunks = splitter.split_documents(docs)

    vectorstore = Chroma.from_documents(
        documents=chunks,
        embedding=get_embeddings(),
        persist_directory=CHROMA_DIR,
        collection_name=collection_name,
    )
    vectorstore.persist()
    return vectorstore


# ─── GET EXISTING VECTORSTORE ─────────────────────────────────────────────────
def get_vectorstore(collection_name: str = "reia_docs"):
    return Chroma(
        persist_directory=CHROMA_DIR,
        embedding_function=get_embeddings(),
        collection_name=collection_name,
    )


# ─── RAG CHAIN ────────────────────────────────────────────────────────────────
def get_rag_chain(collection_name: str = "reia_docs"):
    vectorstore = get_vectorstore(collection_name)
    retriever = vectorstore.as_retriever(
        search_type="similarity",
        search_kwargs={"k": 5},
    )

    chain = (
        {"context": retriever, "question": RunnablePassthrough()}
        | get_llm()
        | StrOutputParser()
    )
    return chain


# ─── LANGCHAIN TOOL (plugs into agent.py) ─────────────────────────────────────
@tool
def analyze_financials(query: str) -> str:
    """Analyze financial data from uploaded documents based on the query."""
    try:
        chain = get_rag_chain()
        return chain.invoke(query)
    except Exception as e:
        return f"RAG error: {str(e)} — make sure a document has been uploaded first."