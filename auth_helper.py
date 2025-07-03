import requests
import config

def get_access_token():
    """
    Obtains an access token using the client credentials grant flow.
    Requires tenant ID, client ID, and client secret to be configured in config.py.
    """
    token_url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': config.CLIENT_ID,
        'client_secret': config.CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    }

    print("Requesting access token...")
    response = requests.post(token_url, data=token_data)
    response.raise_for_status()
    access_token = response.json().get('access_token')
    print("Access token successfully obtained.")
    return access_token