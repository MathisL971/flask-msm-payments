from flask import Flask, jsonify, request
from flask_mail import Mail, Message
from waitress import serve
from flask_cors import CORS
import win32com.client
import pythoncom
import os
import stripe
from datetime import datetime
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
    query = 'SELECT IDFComptabiliteDB, IDFFactureDB, NoFacture, Debit, Credit FROM ' + table
    if request.args:
        query += ' WHERE '
        for i, key in enumerate(request.args):
            query += key + ' = ' + request.args.get(key)
            if i < len(request.args) - 1:
                query += " AND "
            
    results = query_database(query, table)

    return jsonify(results)

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
      send_success_charge_email(charge)
      return jsonify('Email sent!')
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

def send_success_charge_email(charge):
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
  created = datetime.fromtimestamp(charge.created)
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

def send_failed_database_email(payment_intent_id):
  msg = Message(
    'Database Error - Payment Intent ID: ' + str(payment_intent_id),
    recipients=[os.getenv('MAIL_USERNAME')]
  )
   
  msg.html = '<h3>Database Error. Payment Intent ID: ' + str(payment_intent_id) + '</3>'
  msg.html += '<p>Please update the database records manually. You should have received an email with more details about the transaction.</p>'

  with app.app_context():
      mail.send(msg)

@app.route('/api/fetch-payment-intent/<int:invoice_id>', methods=['GET'])
def fetch_payment_intent(invoice_id):
    result = stripe.PaymentIntent.search(
       query="metadata['invoice_id']:'" + str(invoice_id) + "'",
       limit=100
    )

    payment_intents = result['data']
    
    if len(payment_intents) == 0:
      return jsonify(payment_intent=None)
    
    succeeded_payment_intents = [payment_intent for payment_intent in payment_intents if payment_intent.status == 'succeeded']

    if len(succeeded_payment_intents) > 0:
      payment_intent = max(succeeded_payment_intents, key=lambda x: x.created)
    else:
      payment_intent = max(payment_intents, key=lambda x: x.created)

    return jsonify(payment_intent=payment_intent)

@app.route('/api/create-payment-intent', methods=['POST'])
def create_payment_intent():
    debit = request.get_json() 
    payment_intent = stripe.PaymentIntent.create(
       amount=int(''.join(debit['Debit'].split(','))),
       currency="cad",
       payment_method_types=["card"],
       metadata={"invoice_id": debit['NoFacture']},
       receipt_email="lefrancmathis@gmail.com"
    )
    return jsonify(payment_intent=payment_intent)

@app.route('/api/create-credit-entry', methods=['POST'])
def create_credit_entry():
    body = request.get_json()
    payment_intent_id = body['paymentIntentId']

    try:  
      payment_intent = stripe.PaymentIntent.retrieve(payment_intent_id)
      invoice_id = payment_intent.metadata['invoice_id']

      table = 'FComptabiliteDB'
      query = 'SELECT IDFFactureDB, NoFacture, Debit, Code_Client, Nom_Client FROM ' + table + ' WHERE NoFacture = ' + invoice_id
      entries = query_database(query, table)

      if len(entries) == 0:
          print('No entries found for invoice ID ' + invoice_id)
          return jsonify(message='No entries found for invoice ID ' + invoice_id), 404
      
      if len(entries) > 1:
          print('Multiple entries already found for invoice ID ' + invoice_id)

      debit = entries[0]
    
      credit = {
        'IDFFactureDB': debit['IDFFactureDB'],
        'NoFacture': debit['NoFacture'],
        'Credit': '.'.join(debit['Debit'].split(',')),
        'Debit': '0',
        'Nature': 3,
        'IDFModePaiementDB': 2,
        'ModePaiementDB': 2,
        'Titre': 'Credit card',
        'IDFCompte': 1,
        'IDFClient': 1,
        'Code_Client': debit['Code_Client'],
        'Nom_Client': debit['Nom_Client'],
        'DATE': datetime.now().strftime('%Y%m%d')
      }

      statement = f'INSERT INTO {table} ({", ".join(credit.keys())}) VALUES ({", ".join([f"'{value}'" for value in tuple(credit.values())])})'
      execute_database(statement, table)

      query = 'SELECT IDFComptabiliteDB, IDFFactureDB, NoFacture, Debit, Credit FROM ' + table
      query += ' WHERE NoFacture = ' + str(credit['NoFacture']) + ' AND Credit != 0' 
      results = query_database(query, table)

      if len(results) == 0:
          send_failed_database_email(payment_intent_id)
          return jsonify(message='No credit entries found for invoice ID ' + invoice_id), 404

      credit = results[0]

      return jsonify(credit)
    except Exception as e:
      send_failed_database_email(payment_intent_id)
      return jsonify(message=str(e)), 500

if __name__ == '__main__':
    #app.run(debug=True)
    print('Listening on port 5000...')
    serve(app, host='0.0.0.0', port=5000)
    

