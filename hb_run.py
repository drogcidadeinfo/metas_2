import gspread
import time
import logging
import pandas as pd
import calendar
import json
import os
from datetime import date, timedelta
from openpyxl import Workbook
from google.oauth2.service_account import Credentials
from googleapiclient.errors import HttpError

creds_json = os.getenv("GSA_CREDENTIALS")
SOURCE_SHEET_ID = os.getenv("SOURCE_SHEET_ID")
SOURCE_WORKSHEET = "VENDAS_VENDEDOR_HB"
TARGET_WORKSHEET = "CAMPANHA HB"

# ---------------- TARGET SHEETS MAPPING ----------------
def get_target_sheets():
    config_json = os.getenv("TARGET_SHEETS_CONFIG")

    if not config_json:
        raise RuntimeError("TARGET_SHEETS_CONFIG not set")

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

# ---------------- RATE LIMITING CONFIG ----------------
# Google Sheets API limits: 60 read requests per minute per user
# We'll implement exponential backoff for rate limiting
MAX_RETRIES = 5
INITIAL_DELAY = 2  # seconds
MAX_DELAY = 60  # seconds
BATCH_DELAY = 10  # seconds between filials (increased from 2)

# ---------------- HELPERS WITH RETRY LOGIC ----------------
def get_gspread_client():
    creds = Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(creds)

def clear_sheet_range(client, sheet_id, worksheet_name, start_row=9, end_row=57):
    def _clear_range():
        logging.info(f"Clearing rows {start_row}-{end_row} in '{worksheet_name}'")
        ws = get_worksheet(client, sheet_id, worksheet_name)
        
        # Clear only column H range to preserve other data
        ws.batch_clear([f"H{start_row}:H{end_row}"])
    
    return retry_with_backoff(_clear_range)

def retry_with_backoff(func, *args, **kwargs):
    """Execute a function with exponential backoff retry logic."""
    delay = INITIAL_DELAY
    for attempt in range(MAX_RETRIES):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "Quota exceeded" in str(e):
                if attempt < MAX_RETRIES - 1:
                    logging.warning(f"Rate limit hit (attempt {attempt + 1}/{MAX_RETRIES}). Waiting {delay} seconds...")
                    time.sleep(delay)
                    delay = min(delay * 2, MAX_DELAY)  # Exponential backoff
                else:
                    logging.error(f"Max retries exceeded for {func.__name__}")
                    raise
            else:
                raise
        except Exception as e:
            logging.error(f"Error in {func.__name__}: {str(e)}")
            raise

def get_worksheet(client, sheet_id, worksheet_name):
    def _get_worksheet():
        return client.open_by_key(sheet_id).worksheet(worksheet_name)
    
    return retry_with_backoff(_get_worksheet)

def get_spreadsheet(client, sheet_id):
    def _get_spreadsheet():
        return client.open_by_key(sheet_id)
    
    return retry_with_backoff(_get_spreadsheet)

def read_google_sheet(client, sheet_id, worksheet_name):
    def _read_sheet():
        logging.info(f"Reading sheet: {worksheet_name}")
        worksheet = get_worksheet(client, sheet_id, worksheet_name)
        values = worksheet.get_all_values()

        if not values:
            return pd.DataFrame()

        return pd.DataFrame(values[1:], columns=values[0])
    
    return retry_with_backoff(_read_sheet)

def read_source_data(client):
    """Read the consolidated source data from VENDAS_VENDEDOR_HB sheet."""
    def _read_source():
        logging.info(f"Reading source data from {SOURCE_WORKSHEET}")
        worksheet = get_worksheet(client, SOURCE_SHEET_ID, SOURCE_WORKSHEET)
        values = worksheet.get_all_values()
        
        if not values or len(values) < 2:
            logging.warning("Source sheet is empty or has no data rows")
            return {}
        
        # Convert to dictionary with Código as key and Valor Vendas as value
        # Assuming first row is headers: Código, Vendedor, Valor Vendas
        source_data = {}
        for row in values[1:]:  # Skip header row
            if len(row) >= 3:
                codigo = row[0]  # Código is first column
                valor_vendas = row[2]  # Valor Vendas is third column
                
                # Clean the valor_vendas (remove R$ and convert to number)
                if valor_vendas:
                    # Remove currency symbol and dots, replace comma with dot for decimal
                    clean_value = str(valor_vendas).replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
                    try:
                        source_data[codigo] = float(clean_value)
                    except ValueError:
                        logging.warning(f"Could not convert value '{valor_vendas}' for code {codigo}")
                        source_data[codigo] = 0
                else:
                    source_data[codigo] = 0
        
        logging.info(f"Loaded {len(source_data)} records from source")
        return source_data
    
    return retry_with_backoff(_read_source)

def write_df_to_sheet(client, df, sheet_id, worksheet_name, start_row=9):
    def _write_sheet():
        logging.info(
            f"Writing {len(df)} rows to '{worksheet_name}' starting at row {start_row}"
        )

        if df.empty:
            logging.warning("DataFrame is empty. Nothing to write.")
            return

        worksheet = get_worksheet(client, sheet_id, worksheet_name)
        
        # Write in chunks to avoid large API calls
        chunk_size = 100
        for i in range(0, len(df), chunk_size):
            chunk = df.iloc[i:i + chunk_size]
            worksheet.update(
                range_name=f"A{start_row + i}",
                values=chunk.values.tolist(),
                value_input_option="USER_ENTERED"
            )
            if i + chunk_size < len(df):  # Don't sleep after last chunk
                time.sleep(1)  # Small delay between chunks
    
    return retry_with_backoff(_write_sheet)

def update_target_values(client, target_sheet_id, worksheet_name, source_data, filial_number):
    """Update column H with values from source_data based on employee codes."""
    
    def _update_values():
        logging.info(f"Updating values for Filial {filial_number} in {worksheet_name}")
        
        # Get the target worksheet
        worksheet = get_worksheet(client, target_sheet_id, worksheet_name)
        
        # Get all values to find employee codes and their rows
        all_values = worksheet.get_all_values()
        
        if not all_values:
            logging.warning(f"Target worksheet {worksheet_name} is empty")
            return
        
        # Find the header row (where "CÓDIGO" is in column B)
        # In the target sheet, data typically starts at row 9 with headers
        header_row_index = None
        data_start_row = None
        
        for i, row in enumerate(all_values):
            if len(row) >= 2 and row[1] == "CÓDIGO":  # Column B index 1
                header_row_index = i
                data_start_row = i + 1
                break
        
        if header_row_index is None:
            logging.warning(f"Could not find header row with 'CÓDIGO' in {worksheet_name}")
            return
        
        logging.info(f"Found headers at row {header_row_index + 1}, data starts at row {data_start_row + 1}")
        
        # Prepare updates for column H
        updates = []
        updated_count = 0
        not_found_count = 0
        
        # Process each data row starting from data_start_row
        for row_num in range(data_start_row, len(all_values)):
            row = all_values[row_num]
            
            # Check if we have at least column B (Código)
            if len(row) < 2 or not row[1].strip():
                continue
            
            codigo = row[1].strip()  # Código is in column B (index 1)
            
            # Check if this code exists in our source data
            if codigo in source_data:
                valor_hb = source_data[codigo]
                # Format as Brazilian currency
                formatted_value = f"R$ {valor_hb:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                
                # Prepare update for column H (index 7)
                updates.append({
                    'range': f'H{row_num + 1}',
                    'values': [[formatted_value]]
                })
                updated_count += 1
            else:
                # Optional: Clear the cell if code not found
                updates.append({
                    'range': f'H{row_num + 1}',
                    'values': [['']]
                })
                not_found_count += 1
        
        # Apply updates in batches
        if updates:
            logging.info(f"Preparing {len(updates)} updates for Filial {filial_number}")
            
            # Split updates into batches of 50 to avoid API limits
            batch_size = 50
            for i in range(0, len(updates), batch_size):
                batch = updates[i:i + batch_size]
                
                # Prepare batch update request
                batch_data = [{
                    'range': update['range'],
                    'values': update['values']
                } for update in batch]
                
                worksheet.batch_update(batch_data)
                logging.info(f"Applied batch {i//batch_size + 1} with {len(batch)} updates")
                
                # Small delay between batches
                if i + batch_size < len(updates):
                    time.sleep(2)
            
            logging.info(f"Updated {updated_count} rows in Filial {filial_number} (cleared {not_found_count} rows)")
        else:
            logging.info(f"No updates needed for Filial {filial_number}")
    
    return retry_with_backoff(_update_values)

def process_filial(client, filial_target):
    """Process a single filial and update its corresponding target sheet."""
    
    # Get the target sheet ID for this filial
    if filial_target not in TARGET_SHEETS:
        logging.warning(f"Skipping Filial {filial_target} - not in TARGET_SHEETS mapping")
        return
    
    target_sheet_id = TARGET_SHEETS[filial_target]
    
    logging.info(f"Processing Filial {filial_target} -> Target Sheet: {target_sheet_id}")
    
    # Read source data once (but we'll read it in main and pass it to avoid multiple reads)
    # For now, we'll read it here - but ideally read once in main and pass to this function
    
    # We need to pass source_data to this function - let's modify this in main

# ---------------- MAIN ----------------
def main():
    client = get_gspread_client()
    
    # Read source data once for all filials
    logging.info("Reading source data from consolidated sheet...")
    source_data = read_source_data(client)
    
    if not source_data:
        logging.error("No source data found. Exiting.")
        return
    
    # Define filials to process (1-18, but skipping 11)
    filials_to_process = [str(i) for i in range(1, 19) if i != 11]
    
    logging.info(f"Processing {len(filials_to_process)} filials: {filials_to_process}")
    
    successful_filials = []
    failed_filials = []
    
    for index, filial_target in enumerate(filials_to_process, 1):
        try:
            logging.info(f"[{index}/{len(filials_to_process)}] Starting Filial {filial_target}")
            
            # Get target sheet ID
            target_sheet_id = TARGET_SHEETS.get(filial_target)
            if not target_sheet_id:
                logging.warning(f"No target sheet ID for Filial {filial_target}")
                failed_filials.append(filial_target)
                continue
            
            logging.info(f"Processing Filial {filial_target} -> Target Sheet: {target_sheet_id}")
            
            # Clear existing values in column H (rows 9-57)
            clear_sheet_range(client, target_sheet_id, TARGET_WORKSHEET, start_row=9, end_row=57)
            
            # Update with new values based on employee codes
            update_target_values(client, target_sheet_id, TARGET_WORKSHEET, source_data, filial_target)
            
            successful_filials.append(filial_target)
            
            logging.info(f"--- Completed Filial {filial_target} ({index}/{len(filials_to_process)}) ---")
            
            # Add delay between filials (increased for better rate limiting)
            if index < len(filials_to_process):  # Don't wait after last filial
                logging.info(f"Waiting {BATCH_DELAY} seconds before next filial...")
                time.sleep(BATCH_DELAY)
            
        except Exception as e:
            logging.error(f"Error processing Filial {filial_target}: {str(e)}")
            failed_filials.append(filial_target)
            
            # If it's a rate limit error, wait longer before continuing
            if "429" in str(e) or "Quota exceeded" in str(e):
                logging.warning(f"Rate limit error detected. Waiting {MAX_DELAY} seconds before continuing...")
                time.sleep(MAX_DELAY)
            continue
    
    # Summary
    logging.info("=" * 50)
    logging.info("PROCESSING SUMMARY")
    logging.info(f"Total filials: {len(filials_to_process)}")
    logging.info(f"Successful: {len(successful_filials)} - {successful_filials}")
    logging.info(f"Failed: {len(failed_filials)} - {failed_filials}")
    
    if failed_filials:
        logging.warning("Some filials failed. Consider running them individually with more delay.")
        # Optionally retry failed filials
        logging.info("Would you like to retry failed filials? (Implement retry logic if needed)")

if __name__ == "__main__":
    main()
