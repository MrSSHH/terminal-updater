import pandas as pd
import logging
import os

# Setup simple logging to console
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

# --- CONFIGURATION ---
INPUT_PATH = os.path.join('Assets', 'Excels', 'Data.xlsx')
OUTPUT_DIR = os.path.join('Assets', 'Excels', 'toUpdate')
TARGET_VERSION = 808
CHUNK_SIZE = 300

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

try:
    df = pd.read_excel(INPUT_PATH)
except Exception as e:
    logger.error(f"Failed to load Excel file at {INPUT_PATH}: {e}")
    raise SystemExit("Error: Check if Assets/Excels/Data.xlsx exists.")

columns = ['Name', 'Customer', 'UDID', 'Version', 'Serial']
forbidden_customers = []

file_num = 0
rows_to_add = []

for _, row in df.iterrows():
    # 1. Filter Forbidden Customers
    if str(row['Customer']) in forbidden_customers:
        continue
    
    # 2. Filter specific hardware
    if str(row['serial']).startswith('EF500'):
        continue
    
    # 3. Version Check
    try:
        if int(row['ver']) < TARGET_VERSION and int(row['to_ver']) < TARGET_VERSION:
            rows_to_add.append({
                'Name': row['Name'],
                'Customer': row['Customer'], 
                'UDID': row['UDID'],           
                'Version': row['ver'],
                'Serial': row['serial']
            })
    except (ValueError, TypeError):
        continue # Skip rows with invalid version data

    # 4. Save in chunks
    if len(rows_to_add) == CHUNK_SIZE:
        output_file = os.path.join(OUTPUT_DIR, f'output_{file_num}.xlsx')
        pd.DataFrame(rows_to_add, columns=columns).to_excel(output_file, index=False)
        logger.info(f"Created: {output_file}")
        
        rows_to_add = []
        file_num += 1

# Save remaining rows
if rows_to_add:
    output_file = os.path.join(OUTPUT_DIR, f'output_{file_num}.xlsx')
    pd.DataFrame(rows_to_add, columns=columns).to_excel(output_file, index=False)
    logger.info(f"Created final: {output_file}")

logger.info("Filtering complete!")
