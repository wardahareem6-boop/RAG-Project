# DocuMind — Multi-format document RAG

1. Extract this ZIP, open the document_rag folder in VS Code, and open its terminal.
2. Create an environment: `python -m venv .venv`
3. Windows activation: `.venv\Scripts\activate`
   macOS/Linux activation: `source .venv/bin/activate`
4. Install: `python -m pip install -r requirements.txt`
5. Run from this folder: `python -m streamlit run app.py`
6. Enter your OpenAI API key, select files, click Build knowledge workspace, then ask questions.

Python 3.11 or 3.12 is recommended. An OpenAI API key with API billing/model access is required.
A ChatGPT subscription does not provide API credits.

Formats: PDF, DOCX, CSV, PPTX, XLSX, XLS. Legacy PPT is supported by conversion
when LibreOffice is installed on the machine running the app. On Windows the
standard Program Files installation is detected; otherwise put soffice on PATH.
Alternatively open PPT in PowerPoint and Save As PPTX. Do not just rename it.

Features: multiple uploads, dark theme, adjustable chunking and retrieval,
file/page/slide/sheet-row references, saved source excerpts during the session,
clear conversation and transcript download. Changed uploads/settings invalidate
previous answers and require a new index. No index persists after the session.

Limitations: text extraction only, no OCR or image/chart interpretation. DOCX
references are block numbers (Word pagination is not available). Spreadsheets
use first-row headers; row references refer to parsed table rows starting at 2.
Excel formula values rely on cached values saved by Excel. Search retrieves a
subset of rows, so it must not be used for exact whole-sheet totals or exhaustive
analysis. Use pandas or a dedicated analysis tool for those tasks. Follow-ups
should name their subject because each retrieval uses the current question.
Model answers and citations should be checked against the displayed sources.

Privacy: extracted content and questions are sent to OpenAI for embeddings and
answers. Do not deploy publicly without authentication and appropriate limits.
API keys are passed per session rather than stored in process-wide environment.

Validation: Python syntax and local extraction tests are run during preparation.
Live OpenAI calls require your own key and were not exercised here.
