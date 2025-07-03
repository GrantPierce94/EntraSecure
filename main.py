import config
import auth_helper
import graph_service

def main():
    print("--- BEGIN: AUTOMATED GUEST ACCOUNT VISIBILITY REPORT ---")
    
    # Step 1: Acquire access token
    access_token = auth_helper.get_access_token()
    
    # Step 2: Retrieve guest users from Microsoft Graph
    print("Fetching guest user accounts from Entra ID...")
    guest_users = graph_service.get_guest_users(access_token)
    
    if not guest_users:
        print("No guest users found. Exiting.")
        return
    
    print(f"Retrieved {len(guest_users)} guest user records.")
    
    # Step 3: Format data for Excel table
    excel_rows = []
    for user in guest_users:
        created = user.get('createdDateTime', '')
        mail = user.get('mail', '')
        display_name = user.get('displayName', '')
        user_principal = user.get('userPrincipalName', '')
        row = [display_name, user_principal, mail, created.split('T')[0]]
        excel_rows.append(row)

    # Step 4: Clear existing data in Excel table and write new data
    print("Writing data to Excel report...")
    graph_service.clear_excel_table(
        access_token, config.EXCEL_FILE_ID, config.EXCEL_WORKSHEET_NAME, config.EXCEL_TABLE_NAME
    )
    graph_service.add_rows_to_excel_table(
        access_token, config.EXCEL_FILE_ID, config.EXCEL_WORKSHEET_NAME, config.EXCEL_TABLE_NAME, excel_rows
    )
    
    # Step 5: Generate chart in Excel
    num_rows = len(excel_rows)
    source_range = f"'{config.EXCEL_WORKSHEET_NAME}'!A1:D{num_rows + 1}"
    chart_info = graph_service.create_excel_chart(
        access_token, config.EXCEL_FILE_ID, config.EXCEL_CHART_SHEET, source_range
    )
    chart_name = chart_info['name']
    
    # Step 6: Capture chart image
    print("Generating chart image...")
    chart_image = graph_service.get_chart_image(
        access_token, config.EXCEL_FILE_ID, config.EXCEL_CHART_SHEET, chart_name
    )
    
    # Step 7: Send to Microsoft Teams
    print("Sending summary report to Teams...")
    graph_service.send_teams_notification(
        access_token, config.TEAMS_ID, config.TEAMS_CHANNEL_ID, chart_image
    )

    print("--- COMPLETE: REPORT SENT ---")

if __name__ == '__main__':
    main()