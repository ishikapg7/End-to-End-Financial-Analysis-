# 📊 10-K Financial Analysis Chatbot

An AI-powered financial analysis tool built with Claude Opus 4.6 and Streamlit.
Upload 10-K filings for up to 3 public companies and get instant:

- ✅ 3-Statement Analysis (Income Statement, Balance Sheet, Cash Flow)
- ✅ DCF Valuation Model with WACC, terminal value & implied share price
- ✅ Side-by-side company comparison
- ✅ Formatted Excel report (auto-download)
- ✅ Interactive chat — ask follow-up questions about the financials

## How to Run

pip install -r requirements_financial.txt
streamlit run app.py

## Requirements
- Anthropic API key (console.anthropic.com)
- 10-K filings in PDF or TXT format

## Tech Stack
- Claude Opus 4.6 (Anthropic) — financial data extraction & Q&A
- Streamlit — web UI
- pdfplumber — PDF parsing
- openpyxl — Excel export
