from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
import logging
import os

# Configure logging
log_file_path = os.path.join(os.path.dirname(__file__), 'sheet_extraction.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def connect_to_sheets():
    logger.info("Initiating connection to Google Sheets")
    try:
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = Credentials.from_service_account_file(
            'credentials.json', 
            scopes=SCOPES
        )
        service = build('sheets', 'v4', credentials=creds)
        logger.info("Successfully connected to Google Sheets API")
        return service
    except Exception as e:
        logger.error(f"Failed to connect to Google Sheets: {str(e)}")
        raise

def extract_sheet_data():
    try:
        logger.info("Starting data extraction process")
        service = connect_to_sheets()
        
        SPREADSHEET_ID = '15FMeidgU2Dg7Q4JKPkLAdJmQ3IxWCWJXjhCo9UterCE'
        RANGE_NAME = 'A:Z'
        
        logger.info(f"Fetching data from spreadsheet ID: {SPREADSHEET_ID}")
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            logger.warning("No data found in the spreadsheet")
            return None
            
        logger.info(f"Successfully retrieved {len(values)} rows of data")
        
        df = pd.DataFrame(values[1:], columns=values[0])
        logger.info(f"Created DataFrame with shape: {df.shape}")
        
        required_columns = [
            'Email address', 'Tool used', 'Feature Used',
            'Context Awareness', 'Autonomy', 'Experience',
            'Output Quality', 'Overall Rating', 'Unique ID', 'Pod'
        ]
        
        # Verify all required columns exist
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns: {missing_columns}")
            raise ValueError(f"Missing required columns: {missing_columns}")
        
        filtered_df = df[required_columns]
        logger.info(f"Filtered DataFrame created with {len(filtered_df)} rows and {len(required_columns)} columns")
        
        return filtered_df
        
    except Exception as e:
        logger.error(f"Error during data extraction: {str(e)}", exc_info=True)
        return None

if __name__ == "__main__":
    logger.info("Starting script execution")
    data = extract_sheet_data()
    if data is not None:
        logger.info("Data extraction completed successfully")
        print(data.head())
    else:
        logger.error("Failed to extract data")