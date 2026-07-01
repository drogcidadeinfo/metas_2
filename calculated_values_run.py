import gspread
import time
import logging
import pandas as pd
import calendar
import json
import os
from datetime import date, timedelta, datetime
import pytz
from openpyxl import Workbook
from google.oauth2.service_account import Credentials
from googleapiclient.errors import HttpError

creds_json = os.getenv("GSA_CREDENTIALS")
SOURCE_SHEET_ID = os.getenv("SOURCE_SHEET_ID")
SOURCE_WORKSHEET = "calc"
TARGET_WORKSHEET = "RESUMO"

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

# ---------------- FUNCAO SORTING ORDER ----------------
FUNCAO_SORT_ORDER = [
    "GERENTE",
    "SUBGERENTE", 
    "GERENTE FARMACEUTICO",
    "PROMOTOR DE VENDAS",
    "OPERADOR DE CAIXA",
    "OPERADORA DE CAIXA",
    "FARMACEUTICO"
]

# ---------------- RATE LIMITING CONFIG ----------------
# Google Sheets API limits: 60 read requests per minute per user
# We'll implement exponential backoff for rate limiting
MAX_RETRIES = 5
INITIAL_DELAY = 2  # seconds
MAX_DELAY = 60  # seconds
BATCH_DELAY = 10  # seconds between filials (increased from 2)

# ---------------- LOGGING ----------------
logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s - %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

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
        
        # Clear the entire range
        ws.batch_clear([f"A{start_row}:Z{end_row}"])
    
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

def set_spreadsheet_locale_ptbr(client, sheet_id):
    def _set_locale():
        spreadsheet = client.open_by_key(sheet_id)
        requests = [
            {
                "updateSpreadsheetProperties": {
                    "properties": {
                        "locale": "pt_BR"
                    },
                    "fields": "locale"
                }
            }
        ]
        spreadsheet.batch_update({"requests": requests})
    
    return retry_with_backoff(_set_locale)

def get_worksheet(client, sheet_id, worksheet_name):
    def _get_worksheet():
        return client.open_by_key(sheet_id).worksheet(worksheet_name)
    
    return retry_with_backoff(_get_worksheet)

def get_last_updated_datetime():
    """Get current datetime formatted for last update display."""
    utc_now = datetime.now(pytz.UTC)
    brazil_tz = pytz.timezone('America/Sao_Paulo')  # Brazil/East timezone
    brazil_now = utc_now.astimezone(brazil_tz)
    
    return brazil_now.strftime("%d/%m/%Y %H:%M")

def get_spreadsheet(client, sheet_id):
    def _get_spreadsheet():
        return client.open_by_key(sheet_id)
    
    return retry_with_backoff(_get_spreadsheet)

def get_cmv_percentage_by_filial(client, sheet_id, filial):
    def _get_cmv_percentage():
        ws = get_worksheet(client, sheet_id, "VENDAS_FILIAL")
        values = ws.get_all_values()

        df = pd.DataFrame(values[1:], columns=values[0])

        row = df.loc[df["Filial"] == str(filial)]

        if row.empty:
            raise ValueError(f"No VENDAS_FILIAL data for Filial {filial}")

        custo_total = parse_brl_number(row.iloc[0]["Custo Total"])
        faturamento_total = parse_brl_number(row.iloc[0]["Faturamento Total"])

        if faturamento_total == 0:
            return 0.0

        return custo_total / faturamento_total
    
    return retry_with_backoff(_get_cmv_percentage)

def update_percentage_cell(client, sheet_id, worksheet_name, cell, value):
    def _update_cell():
        ws = get_worksheet(client, sheet_id, worksheet_name)
        ws.update(
            range_name=cell,
            values=[[value]],
            value_input_option="USER_ENTERED"
        )
    
    return retry_with_backoff(_update_cell)

def read_google_sheet(client, sheet_id, worksheet_name):
    def _read_sheet():
        logging.info(f"Reading sheet: {worksheet_name}")
        worksheet = get_worksheet(client, sheet_id, worksheet_name)
        values = worksheet.get_all_values()

        if not values:
            return pd.DataFrame()

        return pd.DataFrame(values[1:], columns=values[0])
    
    return retry_with_backoff(_read_sheet)

def sort_df_by_funcao(df):
    """Sort DataFrame by Função in specified order."""
    if df.empty or "Função" not in df.columns:
        return df
    
    # Create a copy to avoid modifying the original
    df_sorted = df.copy()
    
    # Normalize Função values
    df_sorted["Função_clean"] = (
        df_sorted["Função"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace("  ", " ", regex=False)  # Remove double spaces
    )
    
    # Map to sort order (assign numeric value for sorting)
    funcao_order_map = {funcao: i for i, funcao in enumerate(FUNCAO_SORT_ORDER)}
    
    # Create sort column (assign high number for unknown funções)
    df_sorted["_sort_key"] = df_sorted["Função_clean"].apply(
        lambda x: funcao_order_map.get(x, 999)  # Put unknown at the end
    )
    
    # Sort by sort key (primary) and Colaborador (secondary)
    df_sorted = df_sorted.sort_values(
        by=["_sort_key", "Colaborador"]
    ).drop(columns=["Função_clean", "_sort_key"])
    
    return df_sorted

def parse_brl_number(value):
    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    value = str(value).strip()

    # If it has comma, comma is decimal separator
    if "," in value:
        value = value.replace(".", "").replace(",", ".")
    else:
        # No comma → assume dot is thousand separator or integer
        value = value.replace(".", "")

    return float(value)

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

def get_header_dates():
    today = date.today()
    yesterday = today - timedelta(days=1)

    last_day = calendar.monthrange(today.year, today.month)[1]
    end_of_month = date(today.year, today.month, last_day)

    days_remaining = (end_of_month - today).days

    month_name_pt = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO",
        4: "ABRIL", 5: "MAIO", 6: "JUNHO",
        7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO",
        10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
    }

    '''return {
        "month": month_name_pt[today.month],
        "yesterday": today.strftime("%d/%m/%Y"),
        "days_remaining": days_remaining,
    }'''

    return {
            "month": "JUNHO",
            "yesterday": "30/06/2026",
            "days_remaining": "0",
    }

def get_meta_filial_value(client, sheet_id, filial, column_name):
    def _get_meta_value():
        ws = get_worksheet(client, sheet_id, "META_FILIAL")
        values = ws.get_all_values()

        df = pd.DataFrame(values[1:], columns=values[0])

        result = df.loc[df["Filial"] == str(filial), column_name]

        if result.empty:
            raise ValueError(f"No value found for Filial {filial} in column '{column_name}'")

        return result.iloc[0]
    
    return retry_with_backoff(_get_meta_value)

def get_vendas_filial_value(client, sheet_id, filial, column_name):
    def _get_vendas_value():
        ws = get_worksheet(client, sheet_id, "VENDAS_FILIAL")
        values = ws.get_all_values()

        df = pd.DataFrame(values[1:], columns=values[0])

        result = df.loc[df["Filial"] == str(filial), column_name]

        if result.empty:
            raise ValueError(f"No value found for Filial {filial} in column '{column_name}'")

        return result.iloc[0]
    
    return retry_with_backoff(_get_vendas_value)

def get_meta_by_filial(client, sheet_id, filial):
    def _get_meta():
        ws = get_worksheet(client, sheet_id, "META_FILIAL")
        values = ws.get_all_values()

        df = pd.DataFrame(values[1:], columns=values[0])
        meta = df.loc[df["Filial"] == str(filial), "Number"]

        if meta.empty:
            raise ValueError(f"No META found for Filial {filial}")

        return meta.iloc[0]
    
    return retry_with_backoff(_get_meta)

def get_valor_total_by_filial(client, sheet_id, filial):
    def _get_valor_total():
        ws = get_worksheet(client, sheet_id, "VENDAS_FILIAL")
        values = ws.get_all_values()

        df = pd.DataFrame(values[1:], columns=values[0])
        valor = df.loc[df["Filial"] == str(filial), "Faturamento Total"]

        if valor.empty:
            raise ValueError(f"No Valor Total found for Filial {filial}")

        return parse_brl_number(valor.iloc[0])
    
    return retry_with_backoff(_get_valor_total)

def update_header_values(client, sheet_id, worksheet_name, filial):
    def _update_headers():
        ws = get_worksheet(client, sheet_id, worksheet_name)

        dates = get_header_dates()
        raw_meta_str = get_meta_by_filial(client, SOURCE_SHEET_ID, filial)
        raw_meta = parse_brl_number(raw_meta_str)
        valor_total = get_valor_total_by_filial(client, SOURCE_SHEET_ID, filial)
        
        # ----- OBJETIVOS -----
        cmv_raw = get_meta_filial_value(client, SOURCE_SHEET_ID, filial, "CMV")
        hb_raw = get_meta_filial_value(client, SOURCE_SHEET_ID, filial, "HB")
        tkt_raw = get_meta_filial_value(client, SOURCE_SHEET_ID, filial, "TKT MÉDIO")

        hb = parse_brl_number(hb_raw)
        tkt_medio = parse_brl_number(tkt_raw)
        
        hb_day_raw = get_vendas_filial_value(client, SOURCE_SHEET_ID, filial, "Faturamento HB")
        tkt_day_raw = get_vendas_filial_value(client, SOURCE_SHEET_ID, filial, "Ticket Médio")

        hb_day = parse_brl_number(hb_day_raw)
        tkt_day = parse_brl_number(tkt_day_raw)

        # get last updated time
        last_updated = get_last_updated_datetime()

        # Update cells with small delays between them
        updates = [
            ("B2", [[filial]]),
            ("B3", [[dates["month"]]]),
            ("B4", [[dates["yesterday"]]]),
            ("C2:C4", [[dates["days_remaining"]]] * 3),
            ("B5:C5", [[float(raw_meta), float(raw_meta)]]),
            ("D2:D3", [[valor_total], [valor_total]]),
            ("B6", [[last_updated]]),
        ]
        
        for cell, values in updates:
            ws.update(range_name=cell, values=values)
            time.sleep(0.5)  # Small delay between updates
        
        # Update I and J columns
        ws.update(
            range_name="I2:I4",
            values=[
                [cmv_raw],     # CMV (no parsing)
                [hb],          # HB (parsed)
                [tkt_medio],   # TKT MÉDIO (parsed)
            ],
            value_input_option="USER_ENTERED"
        )
        time.sleep(0.5)
        
        ws.update(
            range_name="J3:J4",
            values=[
                [hb_day],          # HB (parsed)
                [tkt_day],   # TKT MÉDIO (parsed)
            ],
            value_input_option="USER_ENTERED"
        )
    
    return retry_with_backoff(_update_headers)

# ---------------- DATA PROCESSING ----------------

def format_qtd_vendas(value):
    try:
        value = float(value)
        if value.is_integer():
            return f"{int(value):,}".replace(",", ".")
        return f"{value:,}".replace(",", ".")
    except Exception:
        return value

def process_excel_data(file_path):
    logging.info("Processing Excel file (vendas vendedor)...")

    df = pd.read_excel(
        file_path,
        header=9,
        dtype={"qtd. vendas": str}
    )

    df.columns = df.columns.str.strip().str.lower()

    current_filial = None
    data = []

    for _, row in df.iterrows():
        codigo_raw = str(row.get("código", "")).strip()

        if "filial:" in codigo_raw.lower():
            current_filial = row.get("unnamed: 3")
            continue

        if codigo_raw.isdigit():
            data.append({
                "Código": codigo_raw,
                "Filial": current_filial,
                "Colaborador": row.get("vendedor"),
                "Qtd Vendas": format_qtd_vendas(row.get("qtd. vendas")),
                "Coluna Vazia": "",
                "Valor Custo": row.get("valor custo"),
                "Faturamento": row.get("valor vendas"),
            })

    result_df = pd.DataFrame(data)
    logging.info(f"Rows processed: {len(result_df)}")

    return result_df

# ---------------- PROCESS SINGLE FILIAL ----------------
'''def process_filial(client, filial_target):
    """Process a single filial and update its corresponding target sheet."""
    
    # Get the target sheet ID for this filial
    if filial_target not in TARGET_SHEETS:
        logging.warning(f"Skipping Filial {filial_target} - not in TARGET_SHEETS mapping")
        return
    
    target_sheet_id = TARGET_SHEETS[filial_target]
    
    logging.info(f"Processing Filial {filial_target} -> Target Sheet: {target_sheet_id}")
    
    # Set locale for target sheet
    set_spreadsheet_locale_ptbr(client, target_sheet_id)
    
    # Add delay after locale setting
    time.sleep(2)

    # Read source data (cache this if reading multiple times)
    df = read_google_sheet(client, SOURCE_SHEET_ID, SOURCE_WORKSHEET)
    logging.info(f"Total rows in source: {len(df)}")
    
    # Add delay after reading source
    time.sleep(3)

    # Filter for current filial
    df_filial = df[df["Filial"] == filial_target]
    logging.info(f"Rows for Filial {filial_target}: {len(df_filial)}")

    # Write data to target sheet
    write_df_to_sheet(
        client,
        df_filial,
        target_sheet_id,
        TARGET_WORKSHEET,
        start_row=9
    )
    
    # Add delay after writing data
    time.sleep(3)

    # Update header values
    update_header_values(
        client,
        target_sheet_id,
        TARGET_WORKSHEET,
        filial_target
    )
    
    # Add delay after updating headers
    time.sleep(2)
    
    # Update CMV percentage
    cmv_percent = get_cmv_percentage_by_filial(
        client,
        SOURCE_SHEET_ID,
        filial_target
    )
    
    # Add delay before final update
    time.sleep(1)

    update_percentage_cell(
        client,
        target_sheet_id,
        TARGET_WORKSHEET,
        "J2",
        cmv_percent
    )
    
    logging.info(f"Successfully processed Filial {filial_target}")'''

def process_filial(client, filial_target):
    """Process a single filial and update its corresponding target sheet."""
    
    # Get the target sheet ID for this filial
    if filial_target not in TARGET_SHEETS:
        logging.warning(f"Skipping Filial {filial_target} - not in TARGET_SHEETS mapping")
        return
    
    target_sheet_id = TARGET_SHEETS[filial_target]
    
    logging.info(f"Processing Filial {filial_target} -> Target Sheet: {target_sheet_id}")
    
    # Set locale for target sheet
    set_spreadsheet_locale_ptbr(client, target_sheet_id)
    
    # Add delay after locale setting
    time.sleep(2)

    # Read source data
    df = read_google_sheet(client, SOURCE_SHEET_ID, SOURCE_WORKSHEET)
    logging.info(f"Total rows in source: {len(df)}")
    
    # Add delay after reading source
    time.sleep(3)

    # Filter for current filial
    df_filial = df[df["Filial"] == filial_target]
    logging.info(f"Rows for Filial {filial_target}: {len(df_filial)}")

    if not df_filial.empty:
        df_filial = sort_df_by_funcao(df_filial)
        logging.info(f"Sorted {len(df_filial)} rows by Função")

    # Clear rows 9-57 in target sheet BEFORE writing new data
    clear_sheet_range(
        client,
        target_sheet_id,
        TARGET_WORKSHEET,
        start_row=9,
        end_row=57
    )
    
    # Add delay after clearing
    time.sleep(1)

    # Write data to target sheet
    write_df_to_sheet(
        client,
        df_filial,
        target_sheet_id,
        TARGET_WORKSHEET,
        start_row=9
    )
    
    # Add delay after writing data
    time.sleep(3)

    # Update header values
    update_header_values(
        client,
        target_sheet_id,
        TARGET_WORKSHEET,
        filial_target
    )
    
    # Add delay after updating headers
    time.sleep(2)
    
    # Update CMV percentage
    cmv_percent = get_cmv_percentage_by_filial(
        client,
        SOURCE_SHEET_ID,
        filial_target
    )
    
    # Add delay before final update
    time.sleep(1)

    update_percentage_cell(
        client,
        target_sheet_id,
        TARGET_WORKSHEET,
        "J2",
        cmv_percent
    )
    
    logging.info(f"Successfully processed Filial {filial_target}")

# ---------------- MAIN ----------------
def main():
    client = get_gspread_client()
    
    # Define filials to process (1-18, but skipping 11)
    filials_to_process = [str(i) for i in range(1, 19) if i != 11]
    
    logging.info(f"Processing {len(filials_to_process)} filials: {filials_to_process}")
    
    successful_filials = []
    failed_filials = []
    
    for index, filial_target in enumerate(filials_to_process, 1):
        try:
            logging.info(f"[{index}/{len(filials_to_process)}] Starting Filial {filial_target}")
            
            process_filial(client, filial_target)
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
