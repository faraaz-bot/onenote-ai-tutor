import os
import requests
import msal

# Configuration - Replace these with your Azure App Registration details
CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "YOUR_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "YOUR_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID", "common")  # 'common' for multi-tenant/personal accounts
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["Notes.ReadWrite", "Notes.ReadWrite.All"]

def get_access_token_interactive():
    """Acquire token interactively (opens browser for login)"""
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )
    
    # Check token cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
            
    # Interactive flow if no cache
    flow = app.init_device_flow(scopes=SCOPE)
    if "user_code" not in flow:
        raise Exception(f"Failed to create device flow. Inner error: {flow}")
        
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Authentication failed: {result.get('error_description')}")

def list_notebooks(access_token):
    """List all OneNote notebooks accessible by the user"""
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get("https://graph.microsoft.com/v1.0/me/onenote/notebooks", headers=headers)
    response.raise_for_status()
    return response.json()

def create_page_in_section(access_token, section_id, title, html_content):
    """Create a new OneNote page inside a specific section using HTML content"""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/xhtml+xml"
    }
    
    # Simple XHTML body required by Microsoft Graph OneNote API
    body = f"""
    <!DOCTYPE html>
    <html>
      <head>
        <title>{title}</title>
      </head>
      <body>
        {html_content}
      </body>
    </html>
    """
    
    url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"
    response = requests.post(url, headers=headers, data=body.encode('utf-8'))
    response.raise_for_status()
    return response.json()

if __name__ == "__main__":
    print("Authenticating with Microsoft Graph...")
    try:
        token = get_access_token_interactive()
        print("Authentication successful!")
        
        print("\nFetching notebooks...")
        notebooks = list_notebooks(token)
        for nb in notebooks.get("value", []):
            print(f"- Notebook: {nb['displayName']} (ID: {nb['id']})")
            
    except Exception as e:
        print(f"Error: {e}")
