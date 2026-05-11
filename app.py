import os
import json
import random
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
import datetime
import msal
from functools import wraps
from dotenv import load_dotenv
from werkzeug.middleware.proxy_fix import ProxyFix

load_dotenv()

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "super-secret-family-key")
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

# Configuration paths
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
ACTIVITIES_FILE = os.path.join(DATA_DIR, 'activities.json')
CONFIG_FILE = os.path.join(DATA_DIR, 'config.json')
CLIENT_SECRETS_FILE = os.path.join(os.path.dirname(__file__), 'client_secret.json')
TOKEN_FILE = os.path.join(DATA_DIR, 'token.json')

SCOPES = ['https://www.googleapis.com/auth/calendar.events', 'https://www.googleapis.com/auth/calendar.readonly']

# MSAL Configuration
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET")
MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "common")
MS_AUTHORITY = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
MS_REDIRECT_PATH = "/getAToken"
MS_SCOPE = ["User.Read"]

def _build_msal_app():
    return msal.ConfidentialClientApplication(
        MS_CLIENT_ID, authority=MS_AUTHORITY,
        client_credential=MS_CLIENT_SECRET)

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("ms_login"))
        return f(*args, **kwargs)
    return decorated_function

@app.route("/ms_login")
def ms_login():
    if not MS_CLIENT_ID:
        # Fallback if MS SSO is not configured yet
        session["user"] = {"preferred_username": "local@dev.local", "name": "Local Admin"}
        return redirect(url_for("index"))
        
    session["flow"] = _build_msal_app().initiate_auth_code_flow(
        MS_SCOPE, redirect_uri=url_for("authorized", _external=True))
    return redirect(session["flow"]["auth_uri"])

@app.route(MS_REDIRECT_PATH)
def authorized():
    try:
        result = _build_msal_app().acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args)
        if "error" in result:
            return f"Login fehlgeschlagen: {result.get('error_description')}"
                
        session["user"] = result.get("id_token_claims")
        
        # Muttertag Special
        user_email = session["user"].get("preferred_username", "")
        if user_email.lower() == "amyloreenbluem@gmail.com":
            session["show_mothers_day"] = True
            
    except Exception as e:
        return f"Fehler bei der Anmeldung: {e}"
    return redirect(url_for("index"))

@app.route("/logout")
def logout():
    session.clear()
    if not MS_CLIENT_ID:
        return redirect(url_for("index"))
    return redirect(
        MS_AUTHORITY + "/oauth2/v2.0/logout" +
        "?post_logout_redirect_uri=" + url_for("index", _external=True))

def get_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {"calendar_id": None}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

def get_activities():
    with open(ACTIVITIES_FILE, 'r') as f:
        return json.load(f)

@app.route('/')
@login_required
def index():
    maps_api_key = os.environ.get("MAPS_API_KEY", "")
    show_mothers_day = session.pop('show_mothers_day', False)
    return render_template('index.html', maps_api_key=maps_api_key, show_mothers_day=show_mothers_day)

@app.route('/get_activity')
@login_required
def get_random_activity():
    activities = get_activities()
    selected = random.choice(activities)
    return jsonify(selected)

@app.route('/admin')
@login_required
def admin():
    config = get_config()
    calendars = []
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except ValueError:
            os.remove(TOKEN_FILE)
            creds = None
    
    if creds and creds.valid:
        try:
            service = build('calendar', 'v3', credentials=creds)
            calendar_list = service.calendarList().list().execute()
            calendars = calendar_list.get('items', [])
        except Exception as e:
            flash(f"Fehler beim Abrufen der Kalender: {e}")
    
    client_secret_exists = os.path.exists(CLIENT_SECRETS_FILE)
    
    return render_template('admin.html', calendars=calendars, config=config, client_secret_exists=client_secret_exists)

@app.route('/authorize')
@login_required
def authorize():
    if not os.path.exists(CLIENT_SECRETS_FILE):
        flash("client_secret.json fehlt! Bitte lade es von Google Cloud Console herunter.")
        return redirect(url_for('admin'))
        
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE, scopes=SCOPES)
    flow.redirect_uri = url_for('oauth2callback', _external=True)
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent')
    session['state'] = state
    return redirect(authorization_url)

@app.route('/oauth2callback')
def oauth2callback():
    state = session['state']
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE, scopes=SCOPES, state=state)
    flow.redirect_uri = url_for('oauth2callback', _external=True)
    
    authorization_response = request.url
    flow.fetch_token(authorization_response=authorization_response)
    
    credentials = flow.credentials
    with open(TOKEN_FILE, 'w') as token_file:
        token_file.write(credentials.to_json())
        
    flash("Google Kalender erfolgreich verbunden!")
    return redirect(url_for('admin'))

@app.route('/save_calendar', methods=['POST'])
@login_required
def save_calendar():
    calendar_id = request.form.get('calendar_id')
    config = get_config()
    config['calendar_id'] = calendar_id
    save_config(config)
    flash("Kalender gespeichert!")
    return redirect(url_for('admin'))

@app.route('/add_event', methods=['POST'])
@login_required
def add_event():
    data = request.json
    config = get_config()
    calendar_id = config.get('calendar_id')
    
    if not calendar_id:
        return jsonify({"success": False, "error": "Kein Kalender ausgewählt. Bitte im Admin-Bereich einstellen."}), 400
        
    if not os.path.exists(TOKEN_FILE):
        return jsonify({"success": False, "error": "Nicht mit Google verbunden. Bitte im Admin-Bereich einstellen."}), 401
        
    try:
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    except ValueError:
        os.remove(TOKEN_FILE)
        return jsonify({"success": False, "error": "Google Anmeldung abgelaufen. Bitte im Admin-Bereich neu verbinden."}), 401
        
    service = build('calendar', 'v3', credentials=creds)
    
    # Calculate times
    if data.get('event_datetime'):
        try:
            start_dt = datetime.datetime.strptime(data.get('event_datetime'), "%Y-%m-%dT%H:%M")
            start_iso = start_dt.isoformat()
            end_iso = (start_dt + datetime.timedelta(hours=3)).isoformat()
        except ValueError:
            now = datetime.datetime.utcnow()
            start_iso = (now + datetime.timedelta(hours=1)).isoformat() + 'Z'
            end_iso = (now + datetime.timedelta(hours=4)).isoformat() + 'Z'
    else:
        now = datetime.datetime.utcnow()
        start_iso = (now + datetime.timedelta(hours=1)).isoformat() + 'Z'
        end_iso = (now + datetime.timedelta(hours=4)).isoformat() + 'Z'
    
    # Check if 'hailey_lars_alone' is True
    title = data.get('title', 'Familienausflug')
    if data.get('hailey_lars_alone'):
        title = "Hailey & Lars: " + title
    
    event = {
        'summary': title,
        'location': data.get('destination', ''),
        'description': data.get('description', ''),
        'start': {
            'dateTime': start_iso,
            'timeZone': 'Europe/Berlin',
        },
        'end': {
            'dateTime': end_iso,
            'timeZone': 'Europe/Berlin',
        },
    }
    
    try:
        event = service.events().insert(calendarId=calendar_id, body=event).execute()
        return jsonify({"success": True, "event_link": event.get('htmlLink')})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    # Insecure transport for local testing of OAuth
    os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
    app.run(debug=True, host='0.0.0.0', port=5000)
