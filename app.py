
import os
import sqlite3
import smtplib
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import Flask, request, jsonify, render_template, send_file

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Initialize SQLite database
def init_db():
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
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
            c.execute("INSERT OR IGNORE INTO rma_sequence (id, last_sequence) VALUES (1, 489)")
            conn.commit()
            logger.info("Database initialized successfully")
    except sqlite3.Error as e:
        logger.error(f"Database initialization failed: {e}")

# Generate sequential RMA token
def generate_rma_token():
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            c = conn.cursor()
            c.execute("SELECT last_sequence FROM rma_sequence WHERE id = 1")
            result = c.fetchone()
            last_sequence = result[0] if result else 489
            new_sequence = last_sequence + 1
            c.execute("UPDATE rma_sequence SET last_sequence = ? WHERE id = 1", (new_sequence,))
            conn.commit()
            logger.info(f"Generated RMA token: MES-RMA-{new_sequence}")
            return f"MES-RMA-{new_sequence}"
    except sqlite3.Error as e:
        logger.error(f"Error generating RMA token: {e}")
        return None

# Send email with RMA token
def send_rma_email(customer_email, issues_observed, device_serial_number, token_no, is_closure=False):
    sender_email = "payal.chaudhari@messung.com"
    sender_password = "fasmvqyajmbavkhr"  # App-specific password
    if is_closure:
        subject = f"RMA Request Closed - Token No: {token_no}"
        body = f"""
Dear Customer,

Your request ticket (#{token_no} - {issues_observed} - Device Sr No: {device_serial_number}) has been deemed closed.

Best regards,
Messung Systems Pvt. Ltd. (Ourican Automation)
"""
    else:
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
        logger.info(f"{'Closure' if is_closure else 'Confirmation'} email sent to {customer_email}")
        return True
    except Exception as e:
        logger.error(f"Email sending failed: {e}")
        return False

@app.route('/')
def index():
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM rma_requests")
            requests = c.fetchall()
            logger.info("Fetched all RMA requests")
        return render_template('index.html', requests=requests, search_results=[])
    except sqlite3.Error as e:
        logger.error(f"Error fetching RMA requests: {e}")
        return render_template('index.html', requests=[], search_results=[], error="Database error")

@app.route('/submit_rma', methods=['POST'])
def submit_rma():
    logger.debug("Received request to /submit_rma")
    try:
        data = request.form
        logger.debug(f"Form data: {dict(data)}")
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
        if not rma:
            logger.error("Failed to generate RMA token")
            return jsonify({'message': 'Failed to generate RMA token', 'success': False}), 500
        faulty_device_status = data.get('faulty_device_status')
        remark = data.get('remark')
        device_status = data.get('device_status', 'Open')  # Default to Open
        r1 = data.get('r1')
        r2 = data.get('r2')
        r3 = data.get('r3')
        customer_email = data.get('customer_email')
        token_no = rma
        db_path = os.getenv('DATABASE_PATH', 'rma.db')
        with sqlite3.connect(db_path, timeout=10) as conn:
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
            logger.debug(f"Inserted RMA: {token_no}")
        email_sent = send_rma_email(customer_email, issues_observed, device_serial_number, token_no)
        return jsonify({
            'message': 'RMA request submitted successfully!',
            'token_no': token_no,
            'email_sent': email_sent,
            'success': True
        })
    except sqlite3.Error as e:
        logger.error(f"Database error in submit_rma: {e}")
        return jsonify({'message': f'Database error: {str(e)}', 'success': False}), 500
    except Exception as e:
        logger.error(f"Error in submit_rma: {e}")
        return jsonify({'message': f'Server error: {str(e)}', 'success': False}), 500

@app.route('/edit_rma/<token>', methods=['GET'])
def edit_rma(token):
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM rma_requests WHERE token_no = ?", (token,))
            rma = c.fetchone()
        if rma:
            rma_data = {
                'month': rma[1], 'date_of_issue': rma[2], 'project': rma[3], 'location': rma[4],
                'si_client': rma[5], 'product': rma[6], 'device_serial_number': rma[7],
                'delivered_material_date': rma[8], 'issues_observed': rma[9], 'emd_observation': rma[10],
                'solutions': rma[11], 'replacement_dc_no': rma[12], 'tested_by_messung_engineer': rma[13],
                'rma': rma[14], 'faulty_device_status': rma[15], 'remark': rma[16], 'device_status': rma[17],
                'r1': rma[18], 'r2': rma[19], 'r3': rma[20], 'token_no': rma[21], 'customer_email': rma[22]
            }
            logger.debug(f"Fetched RMA for editing: {token}")
            return render_template('index.html', requests=[], search_results=[], edit_rma=rma_data)
        else:
            logger.warning(f"RMA token not found: {token}")
            return render_template('index.html', requests=[], search_results=[], error="RMA token not found")
    except sqlite3.Error as e:
        logger.error(f"Error in edit_rma: {e}")
        return render_template('index.html', requests=[], search_results=[], error="Database error")

@app.route('/update_rma/<token>', methods=['POST'])
def update_rma(token):
    logger.debug(f"Received request to /update_rma/{token}")
    try:
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
        db_path = os.getenv('DATABASE_PATH', 'rma.db')
        with sqlite3.connect(db_path, timeout=10) as conn:
            c = conn.cursor()
            c.execute('''UPDATE rma_requests SET
                month = ?, date_of_issue = ?, project = ?, location = ?, si_client = ?, 
                product = ?, device_serial_number = ?, delivered_material_date = ?, 
                issues_observed = ?, emd_observation = ?, solutions = ?, replacement_dc_no = ?, 
                tested_by_messung_engineer = ?, rma = ?, faulty_device_status = ?, 
                remark = ?, device_status = ?, r1 = ?, r2 = ?, r3 = ?, customer_email = ?
                WHERE token_no = ?''',
                      (month, date_of_issue, project, location, si_client, product,
                       device_serial_number, delivered_material_date, issues_observed,
                       emd_observation, solutions, replacement_dc_no, tested_by_messung_engineer,
                       rma, faulty_device_status, remark, device_status, r1, r2, r3,
                       customer_email, token))
            conn.commit()
            logger.info(f"RMA {token} updated successfully")
        return jsonify({'message': 'RMA updated successfully!', 'success': True})
    except sqlite3.Error as e:
        logger.error(f"Database error in update_rma: {e}")
        return jsonify({'message': f'Database error: {str(e)}', 'success': False}), 500
    except Exception as e:
        logger.error(f"Error in update_rma: {e}")
        return jsonify({'message': f'Server error: {str(e)}', 'success': False}), 500

@app.route('/delete_rma/<token>', methods=['POST'])
def delete_rma(token):
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    logger.debug(f"Received request to /delete_rma/{token}")
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            c = conn.cursor()
            c.execute("DELETE FROM rma_requests WHERE token_no = ?", (token,))
            conn.commit()
            logger.info(f"RMA {token} deleted successfully")
        return jsonify({'message': 'RMA deleted successfully!', 'success': True})
    except sqlite3.Error as e:
        logger.error(f"Error in delete_rma: {e}")
        return jsonify({'message': f'Database error: {str(e)}', 'success': False}), 500

@app.route('/close_rma/<token>', methods=['POST'])
def close_rma(token):
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    logger.debug(f"Received request to /close_rma/{token}")
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            c = conn.cursor()
            # Fetch RMA details
            c.execute("SELECT customer_email, issues_observed, device_serial_number FROM rma_requests WHERE token_no = ?", (token,))
            rma = c.fetchone()
            if not rma:
                logger.warning(f"RMA token not found: {token}")
                return jsonify({'message': 'RMA token not found', 'success': False}), 404
            customer_email, issues_observed, device_serial_number = rma
            # Update device_status to Closed
            c.execute("UPDATE rma_requests SET device_status = ? WHERE token_no = ?", ('Closed', token))
            conn.commit()
            logger.info(f"RMA {token} marked as Closed")
        # Send closure email
        email_sent = send_rma_email(customer_email, issues_observed, device_serial_number, token, is_closure=True)
        return jsonify({
            'message': 'RMA closed successfully!' + (' Email sent!' if email_sent else ' Email failed.'),
            'success': True
        })
    except sqlite3.Error as e:
        logger.error(f"Database error in close_rma: {e}")
        return jsonify({'message': f'Database error: {str(e)}', 'success': False}), 500
    except Exception as e:
        logger.error(f"Error in close_rma: {e}")
        return jsonify({'message': f'Server error: {str(e)}', 'success': False}), 500

@app.route('/search', methods=['POST'])
def search_rma():
    search_term = request.form.get('search_term', '').strip()
    search_type = request.form.get('search_type', 'rma')
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    logger.debug(f"Search request: term={search_term}, type={search_type}")
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
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
        with sqlite3.connect(db_path, timeout=10) as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM rma_requests")
            requests = c.fetchall()
        return render_template('index.html', requests=requests, search_results=search_results, 
                              search_term=search_term, search_type=search_type)
    except sqlite3.Error as e:
        logger.error(f"Error in search_rma: {e}")
        return render_template('index.html', requests=[], search_results=[], error="Database error")

@app.route('/export_excel')
def export_excel():
    db_path = os.getenv('DATABASE_PATH', 'rma.db')
    logger.debug("Received request to /export_excel")
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            import pandas as pd
            query = """SELECT token_no, month, date_of_issue, project, location, si_client, product, 
                              device_serial_number AS 'Device Serial Number', delivered_material_date, issues_observed, 
                              emd_observation, solutions, replacement_dc_no, tested_by_messung_engineer, 
                              rma, faulty_device_status, remark, device_status, r1, r2, r3, customer_email 
                       FROM rma_requests"""
            df = pd.read_sql_query(query, conn)
        excel_file = "rma_export.xlsx"
        df.to_excel(excel_file, index=False, engine='openpyxl')
        logger.info("Excel exported successfully")
        return send_file(excel_file, as_attachment=True, download_name="RMA_export.xlsx")
    except Exception as e:
        logger.error(f"Error exporting Excel: {e}")
        return jsonify({'message': f'Error exporting Excel: {str(e)}', 'success': False}), 500

if __name__ == '__main__':
    init_db()
    port = int(os.getenv('PORT', 5000))  # Heroku port
    app.run(host='0.0.0.0', port=port, debug=True)