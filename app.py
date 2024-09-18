from flask import Flask, jsonify, request, send_from_directory
from flask_mail import Mail, Message
from flask_cors import CORS
import win32com.client
import pythoncom
import os
import stripe
from dotenv import load_dotenv
load_dotenv()

app = Flask(__name__, static_folder='static', static_url_path='/')

app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')   # Enter your email here
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')   # Enter your password here
app.config['MAIL_DEFAULT_SENDER'] = ('Me', os.getenv('MAIL_USERNAME'))

mail = Mail(app)

CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)

stripe.api_key = os.getenv('STRIPE_SK_TEST')
endpoint_secret = os.getenv('STRIPE_ENDPOINT_SECRET')

# FOR SELECT QUERIES
def query_database(query, table):
    pythoncom.CoInitialize()

    try:
        conn = win32com.client.Dispatch('ADODB.Connection')

        provider = 'Provider=PCSOFT.HFSQL'
        ds = 'Data Source=localhost:' + os.getenv('HFSQL_PORT')
        db = 'Initial Catalog=' + os.getenv('HFSQL_DB')
        creds = 'User ID=' + os.getenv('HFSQL_USER') + ';Password=' + os.getenv('HFSQL_PASSWORD')
        ex_props = 'Extended Properties="' + 'Password=' + table + ':' + os.getenv('HFSQL_TABLE_PASSWORD') + ';' + 'Cryptage=' + os.getenv('HFSQL_ENCRYPTION') + '"'

        conn.Open(provider + ';' + ds + ';' + db + ';' + creds + ';' + ex_props)

        rs = win32com.client.Dispatch('ADODB.Recordset')

        rs.Open(query, conn)

        results = []
        while not rs.EOF:
            record = {}
            for field in rs.Fields:
                if (field.Name == 'DATE'):                
                    record[field.Name] = str(field).split(' ')[0]
                else:
                    record[field.Name] = field.Value
            results.append(record)
            rs.MoveNext()

        rs.Close()
        conn.Close()

        return results
    finally:
        pythoncom.CoUninitialize()

# FOR INSERT, UPDATE, DELETE QUERIES
def execute_database(query, table):
    pythoncom.CoInitialize()

    try:
        conn = win32com.client.Dispatch('ADODB.Connection')

        provider = 'Provider=PCSOFT.HFSQL'
        ds = 'Data Source=localhost:' + os.getenv('HFSQL_PORT')
        db = 'Initial Catalog=' + os.getenv('HFSQL_DB')
        creds = 'User ID=' + os.getenv('HFSQL_USER') + ';Password=' + os.getenv('HFSQL_PASSWORD')
        ex_props = 'Extended Properties="' + 'Password=' + table + ':' + os.getenv('HFSQL_TABLE_PASSWORD') + ';' + 'Cryptage=' + os.getenv('HFSQL_ENCRYPTION') + '"'

        conn.Open(provider + ';' + ds + ';' + db + ';' + creds + ';' + ex_props)

        conn.Execute(query)

        conn.Close()
    finally:
        pythoncom.CoUninitialize()

#@app.route('/', defaults={'path': ''})
@app.route('/invoices/<int:invoice_id>')
def root(invoice_id):
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/api/entries', methods=['GET'])
def get_entries():
    table = 'FComptabiliteDB'
    query = 'SELECT IDFFactureDB, NoFacture, Date, Debit, Credit, Code_Client, Nom_Client FROM ' + table
    if request.args:
        query += ' WHERE '
        for i, key in enumerate(request.args):
            query += key + ' = ' + request.args.get(key)
            if i < len(request.args) - 1:
                query += " AND "
            
    results = query_database(query, table)

    return jsonify(results)

@app.route('/api/entries', methods=['POST'])
def create_entry():
    table = 'FComptabiliteDB'
    statement = f'INSERT INTO {table} ({', '.join(request.json.keys())}) VALUES ({', '.join([f"'{value}'" for value in tuple(request.json.values())])})'
    execute_database(statement, table)

    query = 'SELECT NoFacture, Debit, Credit FROM ' + table
    query += ' WHERE NoFacture = ' + str(request.json.get('NoFacture')) + ' AND Credit != 0' 
    new_entry = query_database(query, table)

    return jsonify(new_entry)

@app.route('/webhook', methods=['POST'])
def webhook():
    event = None
    payload = request.data
    sig_header = request.headers['STRIPE_SIGNATURE']

    try:
        event = stripe.Webhook.construct_event(
            payload, sig_header, endpoint_secret
        )
    except ValueError as e:
        # Invalid payload
        raise e
    except stripe.error.SignatureVerificationError as e:
        # Invalid signature
        raise e

    if event['type'] == 'payment_intent.payment_failed':
      payment_intent = event['data']['object']

      msg = Message(
        'Payment Intent Failed For ' + str(payment_intent['metadata']['invoice_id']),
        recipients=[os.getenv('MAIL_USERNAME')]
      )

      msg.body = 'A payment intent has failed for invoice ' + str(payment_intent['metadata']['invoice_id'])
      msg.html = '<p>A payment intent has failed for invoice ' + str(payment_intent['metadata']['invoice_id']) + '</p>'

      with app.app_context():
          mail.send(msg)

      return 'Email sent!'
    elif event['type'] == 'payment_intent.succeeded':
      payment_intent = event['data']['object']

      msg = Message(
        'Payment Intent Succeeded For ' + str(payment_intent['metadata']['invoice_id']),
        recipients=[os.getenv('MAIL_USERNAME')]
      )
      msg.body = 'A payment intent has succeeded for invoice ' + str(payment_intent['metadata']['invoice_id'])
      msg.html = '<p>A payment intent has succeeded for invoice ' + str(payment_intent['metadata']['invoice_id']) + '</p>'
      
      with app.app_context():
          mail.send(msg)
      
      return 'Email sent!'
    else:
      print('Unhandled event type {}'.format(event['type']))

    return jsonify(success=True)

if __name__ == '__main__':
    app.run(debug=True)
