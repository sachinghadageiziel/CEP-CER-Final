import os
import requests
import pandas as pd
import json
import re
from PyPDF2 import PdfReader  # <-- added

# ---------------- CONFIG ----------------
API_URL = "http://localhost:7860/api/v1/run/primaryscreen"
API_KEY = os.getenv("LANGFLOW_API_KEY")  # Must be set in env vars # This script works without LANGFLOW_API_KEY because the local LangFlow server
                                         # is not enforcing authentication. If you enable API key security later,
                                         # make sure to set LANGFLOW_API_KEY in your environment before running.



INPUT_EXCEL = "All-Merged.xlsx"   # Excel with abstracts
INPUT_SHEET = "Master"
OUTPUT_EXCEL = "screening_results.xlsx"

IFU_PDF = "ifu.pdf"  # PDF file should be in same folder
# ----------------------------------------


def read_ifu_from_pdf(pdf_path: str) -> str:
    """Extract all text from a PDF file."""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"IFU PDF not found: {pdf_path}")

    reader = PdfReader(pdf_path)
    text = []
    for page in reader.pages:
        text.append(page.extract_text() or "")
    return "\n".join(text).strip()


def clean_json_text(text: str) -> str:
    """Remove ```json fences and extra text before parsing."""
    text = re.sub(r"^```json", "", text.strip(), flags=re.IGNORECASE | re.MULTILINE)
    text = re.sub(r"^```", "", text.strip(), flags=re.MULTILINE)
    return text.strip("` \n\t")


def call_langflow(ifu: str, abstract: str):
    """Call LangFlow API for one abstract."""
    headers = {
        "Content-Type": "application/json",
        "x-api-key": API_KEY,
    }

    payload = {
        "output_type": "chat",
        "input_type": "text",
        "input_value": "",
        "tweaks": {
            "Prompt-3AwNr": {
                "file": ifu,
                "abstract": abstract
            }
        }
    }

    try:
        response = requests.post(API_URL, json=payload, headers=headers)

        # --- DEBUG ---
        print("\n--- API DEBUG ---")
        print("Status:", response.status_code)
        print("Raw Response:", response.text[:300])
        print("--- END DEBUG ---\n")

        response.raise_for_status()
        return response.json()
    except Exception as e:
        return {"Decision": "ERROR", "Rationale": str(e)}


def main():
    # Read IFU text from PDF
    ifu_text = read_ifu_from_pdf(IFU_PDF)
    print(f" Loaded IFU PDF ({len(ifu_text)} characters)")

    # Load Excel
    df = pd.read_excel(INPUT_EXCEL, sheet_name=INPUT_SHEET)

    if "Abstract" not in df.columns:
        raise ValueError("Excel must contain an 'Abstract' column")

    results = []
    for idx, row in df.iterrows():
        abstract = str(row["Abstract"])
        pmid = row.get("PMID", "")

        print(f"Processing PMID={pmid}...")

        result = call_langflow(ifu_text, abstract)

        # Handle LangFlow wrapper
        if "outputs" in result:
            try:
                text_out = result["outputs"][0]["outputs"][0]["results"]["message"]["text"]
                clean_text = clean_json_text(text_out)
                result = json.loads(clean_text)
            except Exception as e:
                result = {"Decision": "ERROR", "Rationale": f"Parse error: {e}"}

        record = {
            "PMID": pmid,
            "Abstract": abstract,
            "Decision": result.get("Decision"),
            "Category": result.get("Category"),
            "ExcludedCriteria": ",".join(result.get("ExcludedCriteria", []))
            if isinstance(result.get("ExcludedCriteria"), list) else result.get("ExcludedCriteria"),
            "Rationale": result.get("Rationale"),
        }
        results.append(record)

    out_df = pd.DataFrame(results)
    out_df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\n Results saved to {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
