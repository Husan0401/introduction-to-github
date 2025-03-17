from flask import Flask, request, render_template, redirect, url_for, session, send_file
import uuid, os, openpyxl
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from twilio.rest import Client
from datetime import datetime
import requests

app = Flask(__name__)
app.secret_key = 'your-secret-key'

# Google API конфигурациясы
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/gmail.send']
creds = Credentials.from_authorized_user_file('token.json', SCOPES) if os.path.exists('token.json') else InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES).run_local_server(port=0)
sheets_service = build('sheets', 'v4', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)
gmail_service = build('gmail', 'v1', credentials=creds)

# Twilio конфигурациясы
twilio_client = Client('your_twilio_sid', 'your_twilio_token')
twilio_phone = 'your_twilio_phone_number'

SPREADSHEET_ID = 'your_spreadsheet_id'
DRIVE_LETTERS_FOLDER = {'invoices': 'invoices_folder_id', 'acts': 'acts_folder_id', 'others': 'others_folder_id'}

@app.route('/')
def home():
    if 'user' not in session:
        return redirect(url_for('login'))
    transactions = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='Transactions!A:F').execute().get('values', [])
    total_amount = sum(float(t[3]) for t in transactions if len(t) > 3)
    return render_template('home.html', total_amount=total_amount)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['username']  # Transparent Login Form-дағы "username" атауына сәйкес
        password = request.form['password']
        result = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='Users!A:B').execute()
        users = result.get('values', [])
        for user in users:
            if user[0] == email and user[1] == password:
                session['user'] = email
                return redirect(url_for('home'))
        return render_template('login.html', error="Қате логин немесе пароль")
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form['username']
        password = request.form['password']
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range='Users!A:B',
            valueInputOption='RAW', body={'values': [[email, password]]}).execute()
        return redirect(url_for('login'))
    return render_template('register.html')

@app.route('/create_letter', methods=['GET', 'POST'])
def create_letter():
    if 'user' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        letter_type = request.form['letter_type']
        text = request.form['text']
        recipient_email = request.form['recipient_email']
        recipient_phone = request.form.get('recipient_phone', '')
        amount = float(request.form.get('amount', 0))
        letter_id = f"LT-{uuid.uuid4()}"
        vat = amount * 0.12
        total_with_vat = amount + vat

        folder_id = DRIVE_LETTERS_FOLDER.get(letter_type, DRIVE_LETTERS_FOLDER['others'])
        content = f"Хат: {letter_id}\nМәтін: {text}\nСома: {amount} тг\nҚҚС: {vat} тг\nЖалпы: {total_with_vat} тг"
        with open(f"{letter_id}.txt", 'w', encoding='utf-8') as f:
            f.write(content)
        drive_service.files().create(body={'name': f'{letter_id}.txt', 'parents': [folder_id]}, media_body=f"{letter_id}.txt").execute()

        from email.mime.text import MIMEText
        msg = MIMEText(content)
        msg['to'] = recipient_email
        msg['from'] = session['user']
        msg['subject'] = f"Хат {letter_id}"
        raw_msg = {'raw': msg.as_string()}
        gmail_service.users().messages().send(userId='me', body=raw_msg).execute()

        if recipient_phone:
            twilio_client.messages.create(to=recipient_phone, from_=twilio_phone, body=f"Хат {letter_id} жіберілді. Сома: {total_with_vat} тг")

        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range='Transactions!A:F',
            valueInputOption='RAW', body={'values': [[letter_id, session['user'], letter_type, amount, vat, datetime.now().strftime('%Y-%m-%d')]]}).execute()

        requests.post('http://1c-api.example.com/import', json={'id': letter_id, 'amount': amount})
        return redirect(url_for('home'))
    return render_template('create_letter.html')

@app.route('/history')
def history():
    if 'user' not in session:
        return redirect(url_for('login'))
    result = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='Transactions!A:F').execute()
    transactions = result.get('values', [])
    return render_template('history.html', transactions=transactions)

@app.route('/calculator', methods=['GET', 'POST'])
def calculator():
    if 'user' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        salary = float(request.form.get('salary', 0))
        tax_rate = float(request.form.get('tax_rate', 12)) / 100
        tax = salary * tax_rate
        net_salary = salary - tax
        return render_template('calculator.html', result={'tax': tax, 'net_salary': net_salary})
    return render_template('calculator.html')

@app.route('/report', methods=['GET', 'POST'])
def report():
    if 'user' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        start_date = request.form['start_date']
        end_date = request.form['end_date']
        transactions = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='Transactions!A:F').execute().get('values', [])
        filtered = [t for t in transactions if len(t) > 5 and start_date <= t[5] <= end_date]

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['ID', 'Жіберуші', 'Түрі', 'Сома', 'ҚҚС', 'Дата'])
        for t in filtered:
            ws.append(t)
        wb.save("report.xlsx")
        return send_file("report.xlsx", as_attachment=True)
    return render_template('report.html')

@app.route('/chat', methods=['GET', 'POST'])
def chat():
    if 'user' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        message = request.form['message']
        bot_response = "Сұрағыңызды қарастырамын!"
        sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range='Chat!A:C',
            valueInputOption='RAW', body={'values': [[session['user'], message, bot_response]]}).execute()
        return redirect(url_for('chat'))
    chat_history = sheets_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='Chat!A:C').execute().get('values', [])
    return render_template('chat.html', chat_history=chat_history)

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
