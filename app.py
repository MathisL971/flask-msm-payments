from flask import Flask, jsonify, request, send_from_directory
from flask_mail import Mail, Message
from waitress import serve
from flask_cors import CORS
import win32com.client
import pythoncom
import os
import stripe
import datetime
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
      msg.html = '<li>A payment intent has failed for invoice ' + str(payment_intent['metadata']['invoice_id']) + '</li>'

      with app.app_context():
          mail.send(msg)

      return 'Email sent!'
    elif event['type'] == 'payment_intent.succeeded':
      payment_intent = event['data']['object']
      return 'Payment Intent Succeeded!'
    elif event['type'] == 'charge.succeeded':
      charge = event['data']['object']

      msg = Message(
        'Invoice ' + str(charge['metadata']['invoice_id']) + ' was settled',
        recipients=[os.getenv('MAIL_USERNAME')]
      )

      msg.html = '<p>Invoice ' + str(charge['metadata']['invoice_id']) + ' was settled. You will find more details about the event below.' + '</p>'

      msg.html += '<h3>Here are the details for the charge:</h3>'
      msg.html += '<ul>'
      msg.html += '<li>Stripe ID: ' + str(charge.id) + '</li>'
      msg.html += '<li>Amount : ' + str(charge.amount / 100.0) + '</li>'
      msg.html += '<li>Currency: ' + str(charge.currency) + '</li>'
      created = datetime.datetime.fromtimestamp(charge.created)
      msg.html += '<li>Date: ' + str(created) + '</li>'
      msg.html += '</ul>'

      msg.html += '<h3>Here are the details for the payment method:</h3>'
      msg.html += '<ul>'
      msg.html += '<li>Stripe ID: ' + str(charge.payment_method) + '</li>'
      msg.html += '<li>Type: ' + str(charge.payment_method_details.type) + '</li>'
      msg.html += '<li>Brand: ' + str(charge.payment_method_details.card.brand) + '</li>'
      msg.html += '<li>Last 4: ' + str(charge.payment_method_details.card.last4) + '</li>'
      msg.html += '<li>Funding: ' + str(charge.payment_method_details.card.funding) + '</li>'
      msg.html += '<li>Country: ' + str(charge.payment_method_details.card.country) + '</li>'
      msg.html += '</ul>'

      msg.html += '<h3>Here are the details for the receipt:</h3>'
      msg.html += '<ul>'
      msg.html += '<li>Receiver: ' + str(charge.receipt_email) + '</li>'
      msg.html += '<li>URL: ' + str(charge.receipt_url) + '</li>'
      msg.html += '</ul>'
      
      with app.app_context():
          mail.send(msg)

      return 'Email sent!'
    elif event['type'] == 'charge.updated':
      charge = event['data']['object']
      return jsonify(charge)
    elif event['type'] == 'charge.captured':
      charge = event['data']['object']
      return jsonify(charge)
    elif event['type'] == 'charge.failed':
      charge = event['data']['object']
      return jsonify(charge)
    elif event['type'] == 'charge.expired':
      charge = event['data']['object']
      return jsonify(charge)
    elif event['type'] == 'charge.pending':
      charge = event['data']['object']
      return jsonify(charge)
    elif event['type'] == 'charge.refunded':
      charge = event['data']['object']
      return jsonify(charge)
    else:
      print('Unhandled event type {}'.format(event['type']))

    return jsonify(success=True)

if __name__ == '__main__':
    # app.run(debug=True)
    print('Listening on port 5000...')
    serve(app, host='0.0.0.0', port=5000)
    

