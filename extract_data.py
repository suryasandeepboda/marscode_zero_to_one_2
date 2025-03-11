from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd

def connect_to_sheets():
    # Define the scope and credentials
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = Credentials.from_service_account_file(
        'credentials.json', 
        scopes=SCOPES
    )
    
    # Create the service
    service = build('sheets', 'v4', credentials=creds)
    return service

def extract_sheet_data():
    try:
        service = connect_to_sheets()
        
        # Source spreadsheet ID (extracted from your URL)
        SPREADSHEET_ID = '15FMeidgU2Dg7Q4JKPkLAdJmQ3IxWCWJXjhCo9UterCE'
        
        # Range of data to read (adjust the range as needed)
        RANGE_NAME = 'A:Z'  # Reading all columns, we'll filter later
        
        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        
        # Get the values from the sheet
        values = result.get('values', [])
        
        if not values:
            print('No data found.')
            return None
            
        # Convert to DataFrame
        df = pd.DataFrame(values[1:], columns=values[0])
        
        # Select required columns
        required_columns = [
            'Email address',
            'Tool used',
            'Feature Used',
            'Context Awareness',
            'Autonomy',
            'Experience',
            'Output Quality',
            'Overall Rating',
            'Unique ID',
            'Pod'
        ]
        
        # Filter only required columns
        filtered_df = df[required_columns]
        
        return filtered_df
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    # Test the extraction
    data = extract_sheet_data()
    if data is not None:
        print("Data extracted successfully!")
        print(data.head())  # Display first few rows