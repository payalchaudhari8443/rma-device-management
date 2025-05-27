import pandas as pd
import sqlite3
import os

# Connect to SQLite database
db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
try:
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
except sqlite3.DatabaseError as e:
    print(f"Error connecting to database: {e}")
    exit(1)

# Create or verify rma_sequence table
c.execute('''CREATE TABLE IF NOT EXISTS rma_sequence (
                id INTEGER PRIMARY KEY,
                last_sequence INTEGER
            )''')
c.execute("INSERT OR IGNORE INTO rma_sequence (id, last_sequence) VALUES (1, 440)")
conn.commit()

# Generate sequential RMA token
def generate_rma_token():
    c.execute("SELECT last_sequence FROM rma_sequence WHERE id = 1")
    result = c.fetchone()
    last_sequence = result[0] if result else 440
    new_sequence = last_sequence + 1
    c.execute("UPDATE rma_sequence SET last_sequence = ? WHERE id = 1", (new_sequence,))
    conn.commit()
    return f"MES-RMA-{new_sequence}"

# Read Excel file
excel_path = os.path.join(os.path.dirname(__file__), 'rma_data.xlsx')
try:
    df = pd.read_excel(excel_path)
except FileNotFoundError:
    print(f"Error: rma_data.xlsx not found at {excel_path}")
    conn.close()
    exit(1)
except PermissionError:
    print(f"Error: Permission denied for rma_data.xlsx. Ensure file is not open or locked.")
    conn.close()
    exit(1)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    conn.close()
    exit(1)

# Insert data into rma_requests
for index, row in df.iterrows():
    token_no = generate_rma_token()
    customer_email = str(row.get('Customer Email', 'client@example.com'))
    try:
        c.execute('''INSERT INTO rma_requests (
            month, date_of_issue, project, location, si_client, product, 
            device_serial_number, delivered_material_date, issues_observed, 
            emd_observation, solutions, replacement_dc_no, tested_by_messung_engineer, 
            rma, faulty_device_status, remark, device_status, r1, r2, r3, 
            token_no, customer_email
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        (
            str(row.get('Month', '')),
            str(row.get('Date of Issue', '')),
            str(row.get('Project', '')),
            str(row.get('Location', '')),
            str(row.get('SI/Client', '')),
            str(row.get('Product', '')),
            str(row.get('Device Serial Number', '')),  # Changed to Device Serial Number
            str(row.get('Delivered Material Date', '')),
            str(row.get('Issues Observed', '')),
            str(row.get('EMD Observation', '')),
            str(row.get('Solutions', '')),
            str(row.get('Replacement DC No', '')),
            str(row.get('Tested By Messung Engineer', '')),
            str(row.get('RMA', '')),
            str(row.get('Faulty Device Status', '')),
            str(row.get('Remark', '')),
            str(row.get('Device Status', '')),
            str(row.get('R1', '')),
            str(row.get('R2', '')),
            str(row.get('R3', '')),
            token_no,
            customer_email
        ))
    except sqlite3.IntegrityError as e:
        print(f"Error inserting row {index}: {e}")
        continue

conn.commit()
conn.close()
print("Excel data imported successfully!")