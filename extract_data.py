# Remove this line
# from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
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
        creds = service_account.Credentials.from_service_account_file(
            '/Users/surya.sandeep.boda/Desktop/Marscode Zero to One 2/credentials.json', 
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
        RANGE_NAME = 'POD 5!A1:CE1000'
        
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
        logger.info(f"Available columns in DataFrame: {list(df.columns)}")
        
        required_columns = [
            'Email Address', 'Tool being used', 'Feature used',
            'Context Awareness', 'Autonomy', 'Experience',
            'Output Quality', 'Overall Rating', 'Unique ID'
        ]
        
        # Verify all required columns exist
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns: {missing_columns}")
            raise ValueError(f"Missing required columns: {missing_columns}")
        
        filtered_df = df[required_columns]
        logger.info(f"Filtered DataFrame created with {len(filtered_df)} rows and {len(required_columns)} columns")
        
        # Convert rating columns to numeric
        rating_columns = ['Context Awareness', 'Autonomy', 'Experience', 'Output Quality', 'Overall Rating']
        for col in rating_columns:
            filtered_df[col] = pd.to_numeric(filtered_df[col], errors='coerce')
            logger.info(f"Converted {col} to numeric values")
            logger.debug(f"{col} values: {filtered_df[col].describe()}")

        # Calculate Mean Rating
        metrics_for_mean = ['Context Awareness', 'Autonomy', 'Experience', 'Output Quality']
        logger.info(f"Calculating mean rating using metrics: {metrics_for_mean}")
        
        filtered_df['Mean Rating'] = filtered_df[metrics_for_mean].mean(axis=1)
        logger.info(f"Mean Rating statistics: \n{filtered_df['Mean Rating'].describe()}")
        
        # Calculate difference
        filtered_df['Difference'] = filtered_df['Mean Rating'] - filtered_df['Overall Rating']
        logger.info(f"Difference statistics: \n{filtered_df['Difference'].describe()}")
        
        # Log counts of Ok vs Not ok results
        result_counts = filtered_df['Result'].value_counts()
        logger.info(f"Result distribution: \n{result_counts}")
        
        # Determine Result status
        filtered_df['Result'] = filtered_df['Difference'].apply(
            lambda x: 'Ok' if -1 <= x <= 1 else 'Not ok'
        )
        logger.info("Added Result status based on difference criteria")
        
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