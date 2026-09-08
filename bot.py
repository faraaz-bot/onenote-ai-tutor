import os
import requests
from flask import Flask, request, render_template
from dotenv import load_dotenv
from ai_pipeline import OneNoteAITutorPipeline

load_dotenv()

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
TENANT_ID = os.getenv("TENANT_ID", "common")
REDIRECT_URI = os.getenv("REDIRECT_URI", "http://localhost:8000/callback")

app = Flask(__name__)
token_store = {}

@app.route("/")
def index():
    if "access_token" not in token_store:
        auth_url = (
            f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize?"
            f"client_id={CLIENT_ID}&response_type=code&redirect_uri={REDIRECT_URI}&"
            f"response_mode=query&scope=Notes.ReadWrite%20Notes.ReadWrite.All%20User.Read%20offline_access"
        )
        return render_template("index.html", authenticated=False, auth_url=auth_url)
    
    headers = {"Authorization": f"Bearer {token_store['access_token']}"}
    try:
        nb_res = requests.get("https://graph.microsoft.com/v1.0/me/onenote/notebooks", headers=headers)
        nb_res.raise_for_status()
        notebooks_raw = nb_res.json().get("value", [])
        
        notebooks_data = []
        for nb in notebooks_raw:
            nb_id = nb["id"]
            nb_name = nb["displayName"]
            sec_res = requests.get(f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{nb_id}/sections", headers=headers)
            sections = []
            if sec_res.status_code == 200:
                for sec in sec_res.json().get("value", []):
                    sections.append({"id": sec["id"], "name": sec["displayName"]})
            notebooks_data.append({"name": nb_name, "sections": sections})
            
        return render_template("index.html", authenticated=True, notebooks=notebooks_data)
    except Exception as e:
        return f"<h3>Error loading notebooks: {e}</h3><p><a href='/'>Retry Login</a></p>"

@app.route("/callback")
def callback():
    code = request.args.get("code")
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code"
    }
    res = requests.post(token_url, data=payload)
    if res.status_code != 200:
        return f"Token exchange failed: {res.text}"
        
    data = res.json()
    token_store["access_token"] = data["access_token"]
    return '<script>window.location.href="/";</script>'

@app.route("/process-section/<section_id>")
def process_section(section_id):
    if "access_token" not in token_store:
        return "Not authenticated. <a href='/'>Login</a>"
        
    pipeline = OneNoteAITutorPipeline(token_store["access_token"], GEMINI_API_KEY)
    results = pipeline.process_lecture_section(section_id)
    
    html = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Processing Complete - OneNote AI Tutor</title>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-950 text-slate-100 min-h-screen flex items-center justify-center p-6">
        <div class="bg-slate-900 border border-slate-800 p-8 rounded-2xl max-w-lg w-full shadow-2xl">
            <div class="inline-block p-3 bg-emerald-500/10 rounded-full text-emerald-400 mb-4">
                <svg class="w-8 h-8" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"></path></svg>
            </div>
            <h2 class="text-2xl font-bold mb-4">AI Processing Complete!</h2>
            <ul class="space-y-2 mb-6 text-sm text-slate-300 max-h-60 overflow-y-auto">
    """
    for r in results:
        html += f"<li class='bg-slate-800/50 p-2.5 rounded-lg border border-slate-700/50'>✨ {r}</li>"
    html += """
            </ul>
            <a href="/" class="block w-full py-3 text-center bg-indigo-600 hover:bg-indigo-500 text-white font-semibold rounded-xl shadow transition-all">Back to Notebooks</a>
        </div>
    </body>
    </html>
    """
    return html

if __name__ == "__main__":
    app.run(port=8000)
