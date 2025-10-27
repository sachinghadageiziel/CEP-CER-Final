import os
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from bs4 import XMLParsedAsHTMLWarning
import warnings

# Suppress XML parser warnings
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

# === CONFIG ===
all_merged_file = "All-Merged.xlsx"   # File containing PMID + DOI
pmid_file = "pmid.xlsx"               # File to update and download from
download_folder = os.path.join(os.getcwd(), "Downloaded Papers")
os.makedirs(download_folder, exist_ok=True)

# === STEP 1: LOAD DATA ===
print("🔹 Loading data...")
all_merged_df = pd.read_excel(all_merged_file)
pmid_df = pd.read_excel(pmid_file)

# Ensure necessary columns exist
for col in ["DOI", "Status"]:
    if col not in pmid_df.columns:
        pmid_df[col] = ""

# Convert PMIDs to string
all_merged_df["PMID"] = all_merged_df["PMID"].astype(str)
pmid_df["PMID"] = pmid_df["PMID"].astype(str)

# === STEP 2: UPDATE DOI FROM ALL-MERGED ===
print("🔹 Updating DOI column in pmid.xlsx...")
pmid_df["DOI"] = pmid_df["PMID"].map(
    all_merged_df.set_index("PMID")["DOI"].to_dict()
)

# Save updated pmid.xlsx
pmid_df.to_excel(pmid_file, index=False)
print("✅ DOI column updated and saved in pmid.xlsx")

# === STEP 3: DOWNLOAD PAPERS ===
def save_file_from_url(url, filepath):
    """Download content and save to file if not empty."""
    try:
        response = requests.get(url, timeout=20)
        if response.status_code == 200 and len(response.content) > 1000:
            with open(filepath, "wb") as f:
                f.write(response.content)
            return True
    except Exception as e:
        print(f"  ⚠️ Error: {e}")
    return False

print("\n🔹 Starting download of papers...")
for idx, row in pmid_df.iterrows():
    pmid = str(row["PMID"])
    doi = str(row["DOI"]) if pd.notna(row["DOI"]) else ""
    print(f"[{idx+1}/{len(pmid_df)}] Processing PMID {pmid}...")

    # File paths
    safe_name = pmid.replace("/", "_")
    pdf_path = os.path.join(download_folder, f"{safe_name}.pdf")
    html_path = os.path.join(download_folder, f"{safe_name}.html")

    # If PDF already exists, mark as Downloaded
    if os.path.exists(pdf_path):
        print("  ⏭️ PDF already exists")
        pmid_df.loc[idx, "Status"] = "Downloaded"
        continue

    # Skip if no DOI
    if not doi:
        print("  ❌ DOI not found.")
        pmid_df.loc[idx, "Status"] = ""  # leave blank
        continue

    downloaded = False

    # 1️⃣ Try DOI PDF link
    pdf_url = f"https://doi.org/{doi}"
    print(f"  ⏳ Trying DOI link: {pdf_url}")
    downloaded = save_file_from_url(pdf_url, pdf_path)

    # 2️⃣ Try PMC fallback if not downloaded
    if not downloaded:
        print("  ⏳ Trying PMC fallback...")
        try:
            pmc_resp = requests.get(
                f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?ids={doi}&format=json"
            )
            pmc_data = pmc_resp.json()
            pmcid = pmc_data['records'][0].get('pmcid')
            if pmcid:
                xml_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pmcid}/"
                downloaded = save_file_from_url(xml_url, html_path)
        except Exception as e:
            print(f"  ⚠️ PMC fallback failed: {e}")

    # ✅ Only mark as Downloaded if actual PDF exists
    if downloaded and os.path.exists(pdf_path):
        pmid_df.loc[idx, "Status"] = "Downloaded-paper"
        print(f"  ✅ Saved PDF for PMID {pmid}")
    else:
        pmid_df.loc[idx, "Status"] = ""  # leave blank
        print(f"  ❌ Could not download PMID {pmid}")

    time.sleep(1)  # polite delay

# === STEP 4: SAVE FINAL RESULTS ===
pmid_df.to_excel(pmid_file, index=False)
print("\n✅ All done! Only entries with actual PDFs marked as 'Downloaded'.")
print(f"📁 Files saved in: {download_folder}")



import pandas as pd
from Bio import Entrez

# === CONFIG ===
Entrez.email = "your_email@example.com"  # Replace with your email (required by NCBI)
excel_file = "pmid.xlsx"
pmid_column = "PMID"

# === STEP 1: READ EXISTING EXCEL ===
print("🔹 Loading existing pmid.xlsx...")
df = pd.read_excel(excel_file)

# Ensure columns exist
if "PMC_ID" not in df.columns:
    df["PMC_ID"] = ""
if "PDF_Link" not in df.columns:
    df["PDF_Link"] = ""

# === STEP 2: PROCESS ONLY BLANK STATUS ROWS ===
blank_rows = df[df["Status"].isna() | (df["Status"].astype(str).str.strip() == "")]
print(f"🔹 Found {len(blank_rows)} entries with blank Status.")

for i, row in blank_rows.iterrows():
    pmid = str(row[pmid_column])
    print(f"[{i+1}/{len(blank_rows)}] Processing PMID {pmid}...")

    try:
        handle = Entrez.elink(dbfrom="pubmed", id=pmid, linkname="pubmed_pmc")
        record = Entrez.read(handle)
        handle.close()

        # Extract PMC ID if available
        pmc_id = record[0]['LinkSetDb'][0]['Link'][0]['Id']
        pdf_url = f"https://pmc.ncbi.nlm.nih.gov/articles/PMC{pmc_id}/pdf/"

        df.loc[i, "PMC_ID"] = f"PMC{pmc_id}"
        df.loc[i, "PDF_Link"] = pdf_url

        print(f"✅ Found PMC{pmc_id}")
    except (IndexError, KeyError):
        print(f"❌ No PMC full-text found for PMID {pmid}")
    except Exception as e:
        print(f"⚠️ Error processing PMID {pmid}: {e}")

# === STEP 3: SAVE BACK TO SAME FILE ===
df.to_excel(excel_file, index=False)
print("\n✅ Done! 'pmid.xlsx' updated with new columns PMC_ID and PDF_Link.")

import os
import time
import requests
import pandas as pd

# === CONFIG ===
excel_file = "pmid.xlsx"
download_folder = os.path.join(os.getcwd(), "Downloaded Papers")
os.makedirs(download_folder, exist_ok=True)

# Browser-like headers for PMC PDFs
headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/123.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.ncbi.nlm.nih.gov/pmc/",
    "Accept": "application/pdf",
    "Connection": "keep-alive",
}

# === STEP 1: LOAD EXISTING EXCEL ===
print("🔹 Loading pmid.xlsx...")
df = pd.read_excel(excel_file)

# Ensure columns exist
if "PDF_Link" not in df.columns:
    df["PDF_Link"] = ""
if "Status" not in df.columns:
    df["Status"] = ""

# === STEP 2: DOWNLOAD PDFs WHERE LINK EXISTS ===
link_rows = df[df["PDF_Link"].notna() & (df["PDF_Link"].str.strip() != "")]
print(f"🔹 Found {len(link_rows)} entries with PDF_Link.")

for i, row in link_rows.iterrows():
    pmid = str(row["PMID"])
    pdf_url = row["PDF_Link"].strip()
    pdf_path = os.path.join(download_folder, f"{pmid}.pdf")

    # Skip if PDF already exists
    if os.path.exists(pdf_path):
        print(f"[{i+1}/{len(link_rows)}] {pmid}: PDF already exists")
        df.loc[i, "Status"] = "Downloaded"
        continue

    try:
        print(f"[{i+1}/{len(link_rows)}] {pmid}: Downloading from {pdf_url} ...")
        response = requests.get(pdf_url, headers=headers, timeout=60)
        response.raise_for_status()

        # Save the PDF
        with open(pdf_path, "wb") as f:
            f.write(response.content)

        df.loc[i, "Status"] = "Downloaded"
        print(f"  ✅ Downloaded successfully.")
    except Exception as e:
        df.loc[i, "Status"] = ""  # leave blank if failed
        print(f"  ❌ Failed to download: {e}")

    time.sleep(1)  # polite delay

# === STEP 3: SAVE BACK TO SAME EXCEL ===
df.to_excel(excel_file, index=False)
print("\n✅ Done! 'pmid.xlsx' updated with downloaded PDFs and Status.")
print(f"📁 PDFs saved in: {download_folder}")
