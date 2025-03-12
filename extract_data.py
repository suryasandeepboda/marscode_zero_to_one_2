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

# Add module docstring
"""Module for handling Google Sheets data extraction and processing."""

def connect_to_sheets():
    """Establish connection to Google Sheets API."""
    logger.info("Initiating connection to Google Sheets")
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_file(
            '/Users/surya.sandeep.boda/Desktop/Marscode Zero to One 2/credentials.json', 
            scopes=scopes
        )
        service = build('sheets', 'v4', credentials=creds)
        logger.info("Successfully connected to Google Sheets API")
        return service
    except Exception as e:
        logger.error(f"Failed to connect to Google Sheets: {str(e)}")
        raise

def write_to_target_sheet(service, data):
    """Write processed data to target Google Sheet with formatting."""
    try:
        target_spreadsheet_id = '1FEqiDqqPfb9YHAWBiqVepmmXj22zNqXNNI7NLGCDVak'
        
        # Clear existing data
        logger.info("Clearing existing data from target sheet")
        service.spreadsheets().values().clear(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            range='Sheet1!A:Z'
        ).execute()

        # Clean and prepare data for writing
        data = data.fillna('')
        numeric_columns = ['Context Awareness', 'Autonomy', 'Experience', 
                         'Output Quality', 'Overall Rating', 'Mean Rating', 'Difference']
        
        # Safely convert numeric columns
        for col in numeric_columns:
            if col in data.columns:
                data[col] = data[col].apply(
                    lambda x: round(float(x), 2) if pd.notnull(x) and str(x).strip() != '' else ''
                )

        # Convert DataFrame to list of lists
        headers = list(data.columns)
        values = [headers] + data.values.tolist()
        
        # Convert all values to strings, handling empty values
        values = [[str(cell) if cell != '' else '' for cell in row] for row in values]
        
        logger.info("Writing data to target sheet")
        # Write the data
        service.spreadsheets().values().update(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            range='Sheet1!A1',
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        # Get the last column letter for the range
        last_column = chr(ord('A') + len(headers) - 1)
        
        # Find the Result column index (0-based)
        result_col_idx = headers.index('Result')
        result_col_letter = chr(ord('A') + result_col_idx)
        
        # Apply conditional formatting
        requests = [{
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{
                        'sheetId': 0,
                        'startColumnIndex': result_col_idx,
                        'endColumnIndex': result_col_idx + 1,
                        'startRowIndex': 1  # Skip header row
                    }],
                    'booleanRule': {
                        'condition': {
                            'type': 'TEXT_EQ',
                            'values': [{'userEnteredValue': 'Ok'}]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.7176,
                                'green': 0.8823,
                                'blue': 0.7176
                            }
                        }
                    }
                }
            }
        },
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{
                        'sheetId': 0,
                        'startColumnIndex': result_col_idx,
                        'endColumnIndex': result_col_idx + 1,
                        'startRowIndex': 1  # Skip header row
                    }],
                    'booleanRule': {
                        'condition': {
                            'type': 'TEXT_EQ',
                            'values': [{'userEnteredValue': 'Not ok'}]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.9568,
                                'green': 0.7176,
                                'blue': 0.7176
                            }
                        }
                    }
                }
            }
        }]

        # Apply the formatting
        service.spreadsheets().batchUpdate(
            spreadsheetId=TARGET_SPREADSHEET_ID,
            body={'requests': requests}
        ).execute()
        
        logger.info("Successfully wrote data and applied formatting to target sheet")
        return True
        
    except Exception as e:
        logger.error(f"Failed to write to target sheet: {str(e)}")
        return False

def extract_sheet_data():
    """Extract and process data from source Google Sheet."""
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
        
        # Create a copy of the filtered DataFrame to avoid SettingWithCopyWarning
        filtered_df = df[required_columns].copy()
        logger.info(f"Filtered DataFrame created with {len(filtered_df)} rows and {len(required_columns)} columns")
        
        # Convert rating columns to numeric using .loc
        rating_columns = ['Context Awareness', 'Autonomy', 'Experience', 'Output Quality', 'Overall Rating']
        for col in rating_columns:
            filtered_df.loc[:, col] = pd.to_numeric(filtered_df[col], errors='coerce')
            logger.info(f"Converted {col} to numeric values")
            logger.debug(f"{col} values: {filtered_df[col].describe()}")

        # Calculate Mean Rating using .loc
        metrics_for_mean = ['Context Awareness', 'Autonomy', 'Experience', 'Output Quality']
        logger.info(f"Calculating mean rating using metrics: {metrics_for_mean}")
        
        filtered_df.loc[:, 'Mean Rating'] = filtered_df[metrics_for_mean].mean(axis=1)
        logger.info(f"Mean Rating statistics: \n{filtered_df['Mean Rating'].describe()}")
        
        # Calculate difference using .loc
        filtered_df.loc[:, 'Difference'] = filtered_df['Mean Rating'] - filtered_df['Overall Rating']
        logger.info(f"Difference statistics: \n{filtered_df['Difference'].describe()}")
        
        # Determine Result status before trying to count it
        filtered_df.loc[:, 'Result'] = filtered_df['Difference'].apply(
            lambda x: 'Ok' if -1 <= x <= 1 else 'Not ok'
        )
        logger.info("Added Result status based on difference criteria")
        
        # Now we can safely count Results
        result_counts = filtered_df['Result'].value_counts()
        logger.info(f"Result distribution: \n{result_counts}")
        
        return filtered_df
        
    except Exception as e:
        logger.error(f"Error during data extraction: {str(e)}", exc_info=True)
        return None

# Update main execution
if __name__ == "__main__":
    logger.info("Starting script execution")
    data = extract_sheet_data()
    if data is not None:
        logger.info("Data extraction completed successfully")
        if write_to_target_sheet(connect_to_sheets(), data):
            logger.info("Process completed successfully")
        else:
            logger.error("Failed to write to target sheet")
    else:
        logger.error("Failed to extract data")