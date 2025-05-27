from flask import Flask, request, jsonify, render_template, send_file
import sqlite3
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import os

app = Flask(__name__)

# Initialize SQLite database
def init_db():
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS rma_requests (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    month TEXT,
                    date_of_issue TEXT,
                    project TEXT,
                    location TEXT,
                    si_client TEXT,
                    product TEXT,
                    device_serial_number TEXT,
                    delivered_material_date TEXT,
                    issues_observed TEXT,
                    emd_observation TEXT,
                    solutions TEXT,
                    replacement_dc_no TEXT,
                    tested_by_messung_engineer TEXT,
                    rma TEXT,
                    faulty_device_status TEXT,
                    remark TEXT,
                    device_status TEXT,
                    r1 TEXT,
                    r2 TEXT,
                    r3 TEXT,
                    token_no TEXT UNIQUE,
                    customer_email TEXT
                )''')
    c.execute('''CREATE TABLE IF NOT EXISTS rma_sequence (
                    id INTEGER PRIMARY KEY,
                    last_sequence INTEGER
                )''')
    c.execute("INSERT OR IGNORE INTO rma_sequence (id, last_sequence) VALUES (1, 440)")
    conn.commit()
    conn.close()

# Generate sequential RMA token
def generate_rma_token():
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT last_sequence FROM rma_sequence WHERE id = 1")
    result = c.fetchone()
    last_sequence = result[0] if result else 440
    new_sequence = last_sequence + 1
    c.execute("UPDATE rma_sequence SET last_sequence = ? WHERE id = 1", (new_sequence,))
    conn.commit()
    conn.close()
    return f"MES-RMA-{new_sequence}"

# Send email with RMA token
def send_rma_email(customer_email, issues_observed, device_serial_number, token_no):
    sender_email = "payal.chaudhari@messung.com"
    sender_password = "fasmvqyajmbavkhr"
    subject = f"RMA Request Confirmation - Token No: {token_no}"
    body = f"""
Dear Customer,

Thank you for submitting an RMA request.
Your request ticket (#{token_no} - {issues_observed} - Device Serial Number: {device_serial_number}) has been raised.
Your Token No is: {token_no}
Please use this token for all communications regarding this repair.

Best regards,
Messung Systems Pvt. Ltd. (Ourican Automation)
"""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = customer_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, customer_email, msg.as_string())
        server.quit()
        print(f"Email sent successfully to {customer_email}")
        return True
    except Exception as e:
        print(f"Email sending failed: {e}")
        return False

@app.route('/')
def index():
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT * FROM rma_requests")
    requests = c.fetchall()
    conn.close()
    return render_template('index.html', requests=requests, search_results=[])

@app.route('/submit_rma', methods=['POST'])
def submit_rma():
    data = request.form
    month = data.get('month')
    date_of_issue = data.get('date_of_issue')
    project = data.get('project')
    location = data.get('location')
    si_client = data.get('si_client')
    product = data.get('product')
    device_serial_number = data.get('device_serial_number')
    delivered_material_date = data.get('delivered_material_date')
    issues_observed = data.get('issues_observed')
    emd_observation = data.get('emd_observation')
    solutions = data.get('solutions')
    replacement_dc_no = data.get('replacement_dc_no')
    tested_by_messung_engineer = data.get('tested_by_messung_engineer')
    rma = generate_rma_token()
    faulty_device_status = data.get('faulty_device_status')
    remark = data.get('remark')
    device_status = data.get('device_status')
    r1 = data.get('r1')
    r2 = data.get('r2')
    r3 = data.get('r3')
    customer_email = data.get('customer_email')
    token_no = rma
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('''INSERT INTO rma_requests (
        month, date_of_issue, project, location, si_client, product, 
        device_serial_number, delivered_material_date, issues_observed, 
        emd_observation, solutions, replacement_dc_no, tested_by_messung_engineer, 
        rma, faulty_device_status, remark, device_status, r1, r2, r3, 
        token_no, customer_email
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (month, date_of_issue, project, location, si_client, product,
               device_serial_number, delivered_material_date, issues_observed,
               emd_observation, solutions, replacement_dc_no, tested_by_messung_engineer,
               rma, faulty_device_status, remark, device_status, r1, r2, r3,
               token_no, customer_email))
    conn.commit()
    conn.close()
    email_sent = send_rma_email(customer_email, issues_observed, device_serial_number, token_no)
    return jsonify({
        'message': 'RMA request submitted successfully!',
        'token_no': token_no,
        'email_sent': email_sent
    })

@app.route('/edit_rma/<token>', methods=['GET'])
def edit_rma(token):
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT * FROM rma_requests WHERE token_no = ?", (token,))
    rma = c.fetchone()
    conn.close()
    if rma:
        rma_data = {
            'month': rma[1], 'date_of_issue': rma[2], 'project': rma[3], 'location': rma[4],
            'si_client': rma[5], 'product': rma[6], 'device_serial_number': rma[7],
            'delivered_material_date': rma[8], 'issues_observed': rma[9], 'emd_observation': rma[10],
            'solutions': rma[11], 'replacement_dc_no': rma[12], 'tested_by_messung_engineer': rma[13],
            'rma': rma[14], 'faulty_device_status': rma[15], 'remark': rma[16], 'device_status': rma[17],
            'r1': rma[18], 'r2': rma[19], 'r3': rma[20], 'token_no': rma[21], 'customer_email': rma[22]
        }
        return render_template('index.html', requests=[], search_results=[], edit_rma=rma_data)
    else:
        return render_template('index.html', requests=[], search_results=[], error="RMA token not found")

@app.route('/update_rma/<token>', methods=['POST'])
def update_rma(token):
    data = request.form
    month = data.get('month')
    date_of_issue = data.get('date_of_issue')
    project = data.get('project')
    location = data.get('location')
    si_client = data.get('si_client')
    product = data.get('product')
    device_serial_number = data.get('device_serial_number')
    delivered_material_date = data.get('delivered_material_date')
    issues_observed = data.get('issues_observed')
    emd_observation = data.get('emd_observation')
    solutions = data.get('solutions')
    replacement_dc_no = data.get('replacement_dc_no')
    tested_by_messung_engineer = data.get('tested_by_messung_engineer')
    rma = data.get('rma')
    faulty_device_status = data.get('faulty_device_status')
    remark = data.get('remark')
    device_status = data.get('device_status')
    r1 = data.get('r1')
    r2 = data.get('r2')
    r3 = data.get('r3')
    customer_email = data.get('customer_email')
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('''UPDATE rma_requests SET
        month = ?, date_of_issue = ?, project = ?, location = ?, si_client = ?, 
        product = ?, device_serial_number = ?, delivered_material_date = ?, 
        issues_observed = ?, emd_observation = ?, solutions = ?, 
        replacement_dc_no = ?, tested_by_messung_engineer = ?, rma = ?, 
        faulty_device_status = ?, remark = ?, device_status = ?, 
        r1 = ?, r2 = ?, r3 = ?, customer_email = ?
        WHERE token_no = ?''',
              (month, date_of_issue, project, location, si_client, product,
               device_serial_number, delivered_material_date, issues_observed,
               emd_observation, solutions, replacement_dc_no, tested_by_messung_engineer,
               rma, faulty_device_status, remark, device_status, r1, r2, r3,
               customer_email, token))
    conn.commit()
    conn.close()
    return jsonify({'message': 'RMA updated successfully!'})

@app.route('/search', methods=['POST'])
def search_rma():
    search_term = request.form.get('search_term', '').strip()
    search_type = request.form.get('search_type', 'rma')
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    query = ""
    params = (f"%{search_term}%",) if search_type != 'rma' else (search_term,)
    if search_type == 'rma':
        query = "SELECT * FROM rma_requests WHERE rma = ?"
    elif search_type == 'device_serial_number':
        query = "SELECT * FROM rma_requests WHERE device_serial_number LIKE ?"
    elif search_type == 'si_client':
        query = "SELECT * FROM rma_requests WHERE si_client LIKE ?"
    c.execute(query, params)
    search_results = c.fetchall()
    conn.close()
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute("SELECT * FROM rma_requests")
    requests = c.fetchall()
    conn.close()
    return render_template('index.html', requests=requests, search_results=search_results, 
                          search_term=search_term, search_type=search_type)

@app.route('/export_excel')
def export_excel():
    db_path = os.path.join(os.path.dirname(__file__), 'rma.db')
    conn = sqlite3.connect(db_path)
    query = """SELECT month, date_of_issue, project, location, si_client, product, 
                      device_serial_number AS 'Device Serial Number', delivered_material_date, issues_observed, 
                      emd_observation, solutions, replacement_dc_no, tested_by_messung_engineer, 
                      rma, faulty_device_status, remark, device_status, r1, r2, r3, 
                      token_no, customer_email 
               FROM rma_requests"""
    df = pd.read_sql_query(query, conn)
    conn.close()
    excel_file = "rma_export.xlsx"
    df.to_excel(excel_file, index=False, engine='openpyxl')
    return send_file(excel_file, as_attachment=True, download_name="rma_export.xlsx")

if __name__ == '__main__':
    init_db()
    app.run(host='127.0.0.1', port=5000, debug=True)