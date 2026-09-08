import os
import requests
from dotenv import load_dotenv, set_key

load_dotenv()

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID", "common")

def get_persisted_access_token():
    """Get a valid access token using the saved refresh token, refreshing it automatically if expired."""
    refresh_token = os.getenv("MICROSOFT_REFRESH_TOKEN")
    if not refresh_token:
        raise Exception("No refresh token found. Please log in once via the web dashboard at http://localhost:8000 to authorize the bot.")
        
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
        "scope": "Notes.ReadWrite Notes.ReadWrite.All User.Read offline_access"
    }
    
    res = requests.post(token_url, data=payload)
    if res.status_code != 200:
        raise Exception(f"Failed to refresh token: {res.text}. Please re-authenticate via web UI.")
        
    data = res.json()
    new_access_token = data["access_token"]
    
    # Save new refresh token if rotated by Microsoft
    if "refresh_token" in data:
        new_refresh_token = data["refresh_token"]
        os.environ["MICROSOFT_REFRESH_TOKEN"] = new_refresh_token
        try:
            set_key(".env", "MICROSOFT_REFRESH_TOKEN", new_refresh_token)
        except Exception:
            pass
            
    return new_access_token
