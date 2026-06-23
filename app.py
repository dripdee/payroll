import os
import smtplib
from email.message import EmailMessage
import pandas as pd
from flask import Flask, render_template, request, jsonify, Response
from xhtml2pdf import pisa
from io import BytesIO
from dotenv import load_dotenv
import threading
import queue
import time
from datetime import datetime

load_dotenv()

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('pdf', exist_ok=True)

# Global queue for logs
log_queue = queue.Queue()

def log_msg(msg):
    log_queue.put(msg)
    print(msg)

def generate_pdf_from_html(html_content, output_path):
    with open(output_path, "w+b") as result_file:
        pisa_status = pisa.CreatePDF(html_content, dest=result_file)
    return not pisa_status.err

def process_payslips(filepath):
    try:
        log_msg("Reading SALARY sheet...")
        df = pd.read_excel(filepath, sheet_name='SALARY', header=1)
        
        smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(os.environ.get("SMTP_PORT", 587))
        smtp_user = os.environ.get("SMTP_USER")
        smtp_password = os.environ.get("SMTP_PASSWORD")

        if not smtp_user or not smtp_password:
            log_msg("Error: SMTP credentials not found in .env. Skipping email sending.")
            has_credentials = False
            smtp_conn = None
        else:
            try:
                log_msg("Connecting to email server...")
                smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
                smtp_conn.starttls()
                smtp_conn.login(smtp_user, smtp_password)
                has_credentials = True
            except Exception as e:
                log_msg(f"Failed to connect to SMTP server: {e}")
                has_credentials = False
                smtp_conn = None

        month_year = datetime.now().strftime("%B %Y").upper()

        count = 0
        for index, row in df.iterrows():
            worker_id = row.get('Worker Id')
            if pd.isna(worker_id):
                break # Reached the end of data

            worker_id = str(worker_id).split('.')[0]
            email = str(row.get('EMAIL ID', '')).strip()

            if email.lower() in ["0", "0.0", "none", "", "nan"]:
                log_msg(f"Skipping worker {worker_id}: No valid email address.")
                continue

            log_msg(f"Processing worker {worker_id}...")

            import math
            def to_float(val):
                try:
                    f = float(val)
                    if math.isnan(f):
                        return 0.0
                    return f
                except (ValueError, TypeError):
                    return 0.0

            # Extract data safely
            # Note: User specified 'Wages' maps to basic, and 'Basic Amount' maps to gross wages
            basic = to_float(row.get('Wages', 0))
            incentive = to_float(row.get('incentive', 0))
            extra = to_float(row.get('EXTRA', 0))
            epf = to_float(row.get('E.P.F.', 0))
            esi = to_float(row.get('E.S.I.C', 0))
            advance = to_float(row.get('ADVANCE', 0))
            canteen = to_float(row.get('CANTEEN', 0))
            fine = to_float(row.get('ABSENT', 0))

            data = {
                'month_year': month_year,
                'worker_id': worker_id,
                'worker_name': row.get('Worker Name', ''),
                'designation': row.get('Category Name', ''),
                'department': row.get('Department Name', ''),
                'email': email,
                'uan': str(row.get('UAN', '')).replace('nan', ''), 
                'pf_no': '',
                'bank': str(row.get('BANK NAME', '')).replace('nan', ''),
                'ifsc': str(row.get('IFSC CODE', '')).replace('nan', ''),
                'account': str(row.get('ACCOUNT NO', '')).replace('nan', ''),
                
                'total_hours': to_float(row.get('Total Hour', 0)),
                'rate': to_float(row.get('Per Hour RATE', 0)),
                'gross_wages': to_float(row.get('Basic Amount', 0)),
                'present_days': to_float(row.get('Present', 0)),
                'lop_days': to_float(row.get('Absent', 0)),
                
                'basic': round(basic, 2),
                'incentive': round(incentive, 2),
                'extra': round(extra, 2),
                
                'epf': round(epf, 2),
                'esi': round(esi, 2),
                'advance': round(advance, 2),
                'canteen': round(canteen, 2),
                'fine': round(fine, 2),
            }

            # Calculate totals to match Excel logic:
            # Total Earnings = BASIC - ADVANCE - CANTEEN + INCENTIVE + EXTRA
            data['total_earnings'] = round(basic - advance - canteen + incentive + extra, 2)
            
            # Total Deductions = EPF + ESI + FINE
            data['total_deductions'] = round(epf + esi + fine, 2)
            
            data['net_salary'] = round(data['total_earnings'] - data['total_deductions'], 2)

            # Render HTML
            with app.app_context():
                html_content = render_template('payslip.html', **data)
            
            pdf_path = os.path.join('pdf', f"Payslip_{worker_id}.pdf")
            success = generate_pdf_from_html(html_content, pdf_path)

            if not success:
                log_msg(f"Error generating PDF for {worker_id}.")
                continue

            if has_credentials and smtp_conn:
                try:
                    msg = EmailMessage()
                    msg['Subject'] = f'Payslip for {month_year} - Worker ID: {worker_id}'
                    msg['From'] = smtp_user
                    msg['To'] = email
                    msg.set_content('Please find your attached payslip for this month.')

                    with open(pdf_path, 'rb') as f:
                        pdf_data = f.read()
                    msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename=os.path.basename(pdf_path))

                    smtp_conn.send_message(msg)
                    log_msg(f"Email sent successfully to {email}.")
                    count += 1
                except Exception as e:
                    log_msg(f"Failed to send email to {email}: {e}")

            time.sleep(0.1) # Small pause to prevent rapid-fire blocking, much faster than 1s

        if has_credentials and smtp_conn:
            try:
                smtp_conn.quit()
            except:
                pass
                
        log_msg(f"Finished processing all {count} valid payslips.")
        log_msg("DONE")
    except Exception as e:
        log_msg(f"Critical Error: {str(e)}")
        log_msg("DONE")


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'No file part'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'No selected file'})
    
    if file and file.filename.endswith('.xlsx'):
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'PAY_SLIP_UPLOADED.xlsx')
        file.save(filepath)
        
        # Clear queue
        while not log_queue.empty():
            log_queue.get()
            
        # Start background thread
        thread = threading.Thread(target=process_payslips, args=(filepath,))
        thread.start()
        
        return jsonify({'status': 'success'})
        
    return jsonify({'status': 'error', 'message': 'Invalid file format. Please upload .xlsx'})

@app.route('/stream')
def stream():
    def event_stream():
        while True:
            try:
                msg = log_queue.get(timeout=20)
                yield f"data: {msg}\n\n"
                if msg == "DONE":
                    break
            except queue.Empty:
                yield ": keep-alive\n\n"
    return Response(event_stream(), mimetype="text/event-stream")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
