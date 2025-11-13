import pandas as pd
from Bio import Entrez
import time

# === CONFIG ===
Entrez.email = "your_email@example.com"   #  Replace with your valid email
excel_file = "pmid.xlsx"

# === STEP 1: Load Excel ===
df = pd.read_excel(excel_file)

# Ensure correct column
if 'PMID' not in df.columns:
    raise ValueError("The Excel file must have a column named 'PMID'")

# Add columns if missing
if 'PMCID' not in df.columns:
    df['PMCID'] = ""
if 'PDF_Link' not in df.columns:
    df['PDF_Link'] = ""

# === STEP 2: Fetch PMCID for each PMID ===
for i, row in df.iterrows():
    pmid = str(row['PMID'])

    # Skip already-filled rows
    if pd.notna(row.get('PMCID')) and str(row['PMCID']).startswith('PMC'):
        print(f"[{pmid}] ⏩ Already has PMCID, skipping")
        continue

    try:
        handle = Entrez.elink(dbfrom="pubmed", db="pmc", id=pmid)
        record = Entrez.read(handle)
        handle.close()

        if record and record[0].get("LinkSetDb"):
            pmcid = record[0]["LinkSetDb"][0]["Link"][0]["Id"]
            df.at[i, 'PMCID'] = f"PMC{pmcid}"
            df.at[i, 'PDF_Link'] = f"https://pmc.ncbi.nlm.nih.gov/articles/PMC{pmcid}/pdf/"
            print(f"[{pmid}]  PMCID: PMC{pmcid}")
        else:
            print(f"[{pmid}]  No PMCID found")

    except Exception as e:
        print(f"[{pmid}]  Error: {e}")

    # Respect NCBI API rate limits
    time.sleep(0.5)

# === STEP 3: Save updated Excel ===
df.to_excel(excel_file, index=False)
print(f"\n Updated successfully — 'PMCID' and 'PDF_Link' columns saved to {excel_file}")


import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# === CONFIG ===
excel_file = "pmid.xlsx"
download_folder = os.path.join(os.getcwd(), "pdf download")
os.makedirs(download_folder, exist_ok=True)

# === STEP 1: Load Excel ===
df = pd.read_excel(excel_file)

if 'PDF_Link' not in df.columns:
    raise ValueError("Excel file must have a column named 'PDF_Link'")

if 'Status' not in df.columns:
    df['Status'] = ""

valid_rows = df[df['PDF_Link'].notna() & (df['PDF_Link'] != "")]
print(f" Found {len(valid_rows)} valid PDF links to download.\n")

# === STEP 2: Setup Chrome ===
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
prefs = {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.maximize_window()

# === STEP 3: Download PDFs and update status ===
for i, row in valid_rows.iterrows():
    pmid = str(row.get("PMID", "Unknown"))
    pdf_url = str(row["PDF_Link"]).strip()

    if not pdf_url.startswith("http"):
        print(f"[{pmid}]  Invalid URL, skipping")
        continue

    print(f"[{pmid}]  Opening {pdf_url}")
    try:
        driver.get(pdf_url)
        time.sleep(2)

        # Wait for a new PDF to appear and finish downloading
        timeout = 60  # wait max 60 seconds
        start_time = time.time()
        downloaded_file = None

        while time.time() - start_time < timeout:
            files = [f for f in os.listdir(download_folder) if f.lower().endswith(".pdf")]
            # Only pick files that are fully downloaded (no .crdownload)
            ready_files = [f for f in files if not f.endswith(".crdownload")]
            if ready_files:
                downloaded_file = max(
                    [os.path.join(download_folder, f) for f in ready_files],
                    key=os.path.getctime
                )
                break
            time.sleep(1)

        if not downloaded_file:
            print(f"[{pmid}]  No fully downloaded PDF detected after {timeout}s")
            continue

        # Rename the downloaded file safely
        new_path = os.path.join(download_folder, f"{pmid}.pdf")
        try:
            os.rename(downloaded_file, new_path)
        except PermissionError:
            # Wait a bit and retry
            time.sleep(2)
            os.rename(downloaded_file, new_path)

        # Update status in Excel
        df.at[i, 'Status'] = "Downloaded-S"
        print(f"[{pmid}]  Saved as {pmid}.pdf and status updated")

    except WebDriverException as e:
        print(f"[{pmid}]  Error: {e}")

# === STEP 4: Save Excel ===
df.to_excel(excel_file, index=False)
driver.quit()
print(f"\n All downloads complete. Status column updated in '{excel_file}'")


