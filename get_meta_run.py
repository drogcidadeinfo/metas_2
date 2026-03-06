import time
import logging
import pandas as pd
import gspread
import json
import os
from google.oauth2.service_account import Credentials

# ---------------- CONFIG ----------------
# Get credentials from environment variable
CREDS_JSON = os.getenv("GSA_CREDENTIALS")
if not CREDS_JSON:
    raise RuntimeError("GSA_CREDENTIALS environment variable not set")

SOURCE_SHEET_ID = os.getenv("SOURCE_SHEET_ID")
if not SOURCE_SHEET_ID:
    raise RuntimeError("SOURCE_SHEET_ID environment variable not set")

SOURCE_WORKSHEET = "calc"
TARGET_WORKSHEET = "RESUMO"

# ---------------- TARGET SHEETS MAPPING ----------------
def get_target_sheets():
    config_json = os.getenv("TARGET_SHEETS_CONFIG")
    
    if not config_json:
        raise RuntimeError("TARGET_SHEETS_CONFIG environment variable not set")
    
    try:
        data = json.loads(config_json)
    except json.JSONDecodeError as e:
        raise RuntimeError("Invalid JSON in TARGET_SHEETS_CONFIG") from e
    
    if not isinstance(data, dict):
        raise RuntimeError("TARGET_SHEETS_CONFIG must be a JSON object")
    
    return data

TARGET_SHEETS = get_target_sheets()

if not TARGET_SHEETS:
    raise RuntimeError("TARGET_SHEETS mapping is empty. Check TARGET_SHEETS_CONFIG secret.")

START_ROW = 9  # data starts at row 9
IDENTIFICADOR_COL = 0  # column A
META_COL = 4           # column E  ✅ confirmed

# Rate limiting config
MAX_RETRIES = 5
INITIAL_DELAY = 2  # seconds
MAX_DELAY = 60  # seconds
BATCH_DELAY = 10  # seconds between filials

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
)

# ---------------- AUTH ----------------
def get_client():
    creds = Credentials.from_service_account_info(
        json.loads(CREDS_JSON),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(creds)

# ---------------- RETRY LOGIC ----------------
def retry_with_backoff(func, *args, **kwargs):
    """Execute a function with exponential backoff retry logic."""
    delay = INITIAL_DELAY
    for attempt in range(MAX_RETRIES):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "Quota exceeded" in str(e) or "500" in str(e):
                if attempt < MAX_RETRIES - 1:
                    logging.warning(f"API error (attempt {attempt + 1}/{MAX_RETRIES}). Waiting {delay} seconds...")
                    time.sleep(delay)
                    delay = min(delay * 2, MAX_DELAY)  # Exponential backoff
                else:
                    logging.error(f"Max retries exceeded")
                    raise
            else:
                raise
        except Exception as e:
            logging.error(f"Error: {str(e)}")
            raise

def get_worksheet(client, sheet_id, worksheet_name):
    def _get_worksheet():
        return client.open_by_key(sheet_id).worksheet(worksheet_name)
    
    return retry_with_backoff(_get_worksheet)

# ---------------- HELPERS ----------------
def parse_brl_number(value):
    if value is None:
        return None

    value = str(value).replace("R$", "").strip()
    if value == "":
        return None

    value = value.replace(".", "").replace(",", ".")
    
    try:
        return float(value)
    except ValueError:
        logging.warning(f"Could not parse value: {value}")
        return None

# ---------------- READ META FROM TARGET ----------------
def read_meta_from_target(client, sheet_id, filial_number):
    """Read Meta values from a target sheet with retry logic."""
    
    def _read_meta():
        logging.info(f"Reading Filial {filial_number} - getting worksheet...")
        ws = get_worksheet(client, sheet_id, TARGET_WORKSHEET)
        
        logging.info(f"Reading values from Filial {filial_number}...")
        values = ws.get_all_values()

        meta_map = {}
        rows_processed = 0

        for i, row in enumerate(values[START_ROW - 1:], start=START_ROW):
            if len(row) <= META_COL:
                continue

            identificador = row[IDENTIFICADOR_COL].strip()
            meta_raw = row[META_COL].strip()

            if not identificador or not meta_raw:
                continue

            meta_value = parse_brl_number(meta_raw)
            if meta_value is not None:
                meta_map[identificador] = meta_value
                rows_processed += 1

        logging.info(f"Filial {filial_number}: found {rows_processed} Meta entries")
        return meta_map
    
    return retry_with_backoff(_read_meta)

# ---------------- UPDATE SOURCE SHEET ----------------
def update_source_meta(client, meta_by_id):
    """Update source sheet with collected Meta values using batch updates."""
    
    def _update_source():
        logging.info(f"Opening source sheet: {SOURCE_SHEET_ID}")
        ws = get_worksheet(client, SOURCE_SHEET_ID, SOURCE_WORKSHEET)

        logging.info("Reading source sheet values...")
        values = ws.get_all_values()
        
        if not values:
            logging.warning("Source sheet is empty")
            return

        headers = values[0]
        df = pd.DataFrame(values[1:], columns=headers)

        if "ID" not in df.columns or "Meta" not in df.columns:
            raise ValueError("Source sheet must contain 'ID' and 'Meta' columns")

        meta_col_index = headers.index("Meta") + 1  # 1-indexed for Sheets
        updates = []
        skipped = 0

        logging.info(f"Processing {len(df)} rows in source sheet...")

        for idx, row in df.iterrows():
            row_id = str(row["ID"]).strip()

            if row_id in meta_by_id:
                cell_row = idx + 2  # header offset (1-indexed + skip header)
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(cell_row, meta_col_index),
                    "values": [[meta_by_id[row_id]]],
                })
            else:
                skipped += 1

        logging.info(f"Prepared {len(updates)} updates (skipped {skipped} rows with no matching ID)")

        # Apply updates in batches
        if updates:
            batch_size = 50  # Smaller batch size for reliability
            for i in range(0, len(updates), batch_size):
                batch = updates[i:i + batch_size]
                
                logging.info(f"Applying batch {i//batch_size + 1} with {len(batch)} updates...")
                ws.batch_update(batch)
                
                if i + batch_size < len(updates):
                    time.sleep(2)  # Delay between batches
        
        logging.info(f"✅ Successfully updated {len(updates)} Meta values in source sheet")
    
    return retry_with_backoff(_update_source)

# ---------------- MAIN ----------------
def main():
    logging.info("Starting reverse Meta sync process...")
    
    # Initialize client
    client = get_client()
    logging.info("Google Sheets client initialized successfully")

    combined_meta = {}
    successful_filials = []
    failed_filials = []

    logging.info(f"Processing {len(TARGET_SHEETS)} filials...")

    for idx, (filial, sheet_id) in enumerate(TARGET_SHEETS.items(), 1):
        try:
            logging.info(f"[{idx}/{len(TARGET_SHEETS)}] Reading Filial {filial}")
            
            meta_map = read_meta_from_target(client, sheet_id, filial)
            combined_meta.update(meta_map)
            
            successful_filials.append(filial)
            
            # Add delay between filials
            if idx < len(TARGET_SHEETS):
                logging.info(f"Waiting {BATCH_DELAY} seconds before next filial...")
                time.sleep(BATCH_DELAY)
                
        except Exception as e:
            logging.error(f"Error processing Filial {filial}: {str(e)}")
            failed_filials.append(filial)
            
            # If it's a rate limit error, wait longer
            if "429" in str(e) or "Quota exceeded" in str(e):
                logging.warning(f"Rate limit error. Waiting {MAX_DELAY} seconds...")
                time.sleep(MAX_DELAY)
            continue

    # Summary of reading phase
    logging.info("=" * 50)
    logging.info("READING PHASE SUMMARY")
    logging.info(f"Total filials: {len(TARGET_SHEETS)}")
    logging.info(f"Successfully read: {len(successful_filials)} - {successful_filials}")
    logging.info(f"Failed: {len(failed_filials)} - {failed_filials}")
    logging.info(f"Total Meta entries collected: {len(combined_meta)}")
    logging.info("=" * 50)

    if not combined_meta:
        logging.error("No Meta values collected. Exiting.")
        return

    # Update source sheet
    logging.info("Starting source sheet update phase...")
    
    try:
        update_source_meta(client, combined_meta)
        logging.info("✅ Reverse Meta sync completed successfully")
    except Exception as e:
        logging.error(f"Failed to update source sheet: {str(e)}")
        raise

if __name__ == "__main__":
    main()
