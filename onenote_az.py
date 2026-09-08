import time
import requests

# Public client ID for Microsoft Graph Explorer / third-party apps supporting personal accounts
CLIENT_ID = "de8bc8b5-d9f9-48b1-a8ad-b785da726818"
TENANT = "common"
SCOPES = "Notes.ReadWrite Notes.ReadWrite.All User.Read offline_access"

def device_login():
    dc_url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/devicecode"
    data = {
        "client_id": CLIENT_ID,
        "scope": SCOPES
    }
    res = requests.post(dc_url, data=data)
    res.raise_for_status()
    flow = res.json()
    
    print("=" * 70)
    print(flow["message"])
    print("=" * 70)
    
    token_url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
    payload = {
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "client_id": CLIENT_ID,
        "device_code": flow["device_code"]
    }
    
    interval = flow.get("interval", 5)
    expiry = time.time() + flow.get("expires_in", 900)
    
    while time.time() < expiry:
        time.sleep(interval)
        token_res = requests.post(token_url, data=payload)
        token_data = token_res.json()
        
        if "access_token" in token_data:
            print("\nAuthentication successful!")
            return token_data["access_token"]
            
        error = token_data.get("error")
        if error == "authorization_pending":
            continue
        elif error in ("slow_down", "expired_token", "access_denied"):
            print(f"\nAuthentication error: {error} - {token_data.get('error_description')}")
            break
        else:
            if error:
                print(f"Waiting status: {error} - {token_data.get('error_description', '')}")
            
    raise Exception("Authentication timed out or failed.")

def list_notebooks(token):
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get("https://graph.microsoft.com/v1.0/me/onenote/notebooks", headers=headers)
    res.raise_for_status()
    return res.json()

if __name__ == "__main__":
    token = device_login()
    print("\nFetching OneNote notebooks...")
    notebooks = list_notebooks(token)
    for nb in notebooks.get("value", []):
        print(f"- Notebook: {nb['displayName']} (ID: {nb['id']})")
