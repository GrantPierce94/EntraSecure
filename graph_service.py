import requests

BASE_URL = "https://graph.microsoft.com/v1.0"

def get_guest_users(token):
    """
    Fetches all Entra ID guest users.
    Returns a list of user objects with userType 'Guest'.
    """
    headers = {'Authorization': f'Bearer {token}'}
    users = []
    url = f"{BASE_URL}/users?$filter=userType eq 'Guest'&$top=999"

    while url:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()
        users.extend(data.get('value', []))
        url = data.get('@odata.nextLink', None)

    return users

def clear_excel_table(token, file_id, worksheet, table):
    """
    Clears all rows from a specified Excel table.
    If the table is already empty, skips error.
    """
    headers = {'Authorization': f'Bearer {token}'}
    url = f"{BASE_URL}/me/drive/items/{file_id}/workbook/worksheets/{worksheet}/tables/{table}/rows/clear"
    response = requests.post(url, headers=headers)
    if response.status_code not in [200, 204, 404]:
        response.raise_for_status()
    print(f"Cleared existing data in Excel table: {table}")

def add_rows_to_excel_table(token, file_id, worksheet, table, data_rows):
    """
    Appends rows to an Excel table in the given workbook.
    Data format: list of lists
    """
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f"{BASE_URL}/me/drive/items/{file_id}/workbook/worksheets/{worksheet}/tables/{table}/rows/add"
    payload = {'values': data_rows}
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    print(f"Inserted {len(data_rows)} rows into Excel table: {table}")
    return response.json()

def create_excel_chart(token, file_id, worksheet, source_address):
    """
    Creates a clustered column chart from the given data range.
    Returns the chart metadata.
    """
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f"{BASE_URL}/me/drive/items/{file_id}/workbook/worksheets/{worksheet}/charts/add"
    payload = {
        "type": "ColumnClustered",
        "sourceData": source_address,
        "seriesBy": "Auto"
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    print(f"Chart created on worksheet: {worksheet}")
    return response.json()

def get_chart_image(token, file_id, chart_sheet, chart_name):
    """
    Retrieves a base64-encoded image of the specified chart.
    """
    headers = {'Authorization': f'Bearer {token}'}
    url = f"{BASE_URL}/me/drive/items/{file_id}/workbook/worksheets/{chart_sheet}/charts/{chart_name}/image(width=600,height=400)"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()['value']

def send_teams_notification(token, team_id, channel_id, chart_image_base64):
    """
    Sends a message with the chart image to a Microsoft Teams channel.
    """
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f"{BASE_URL}/teams/{team_id}/channels/{channel_id}/messages"
    message = f"""
        <h2>Automated Guest User Report</h2>
        <p>Below is the current guest user activity and age distribution.</p>
        <img src='data:image/png;base64,{chart_image_base64}' alt='Guest User Chart' />
    """
    payload = {
        "body": {
            "contentType": "html",
            "content": message
        }
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    print("Report posted to Microsoft Teams.")