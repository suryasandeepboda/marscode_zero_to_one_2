import unittest
from unittest.mock import patch, MagicMock
import logging
from google.oauth2 import service_account
from extract_data import connect_to_sheets

class TestConnectToSheets(unittest.TestCase):
    def setUp(self):
        # Configure logging for tests
        logging.basicConfig(level=logging.INFO)
        self.credentials_path = '/Users/surya.sandeep.boda/Desktop/Marscode Zero to One 2/credentials.json'
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets']

    @patch('extract_data.service_account.Credentials')
    @patch('extract_data.build')
    def test_successful_connection(self, mock_build, mock_credentials):
        # Arrange
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        mock_creds = MagicMock()
        mock_credentials.from_service_account_file.return_value = mock_creds

        # Act
        result = connect_to_sheets()

        # Assert
        mock_credentials.from_service_account_file.assert_called_once_with(
            self.credentials_path,
            scopes=self.scopes
        )
        mock_build.assert_called_once_with('sheets', 'v4', credentials=mock_creds)
        self.assertEqual(result, mock_service)

    @patch('extract_data.service_account.Credentials')
    def test_credentials_file_error(self, mock_credentials):
        # Arrange
        mock_credentials.from_service_account_file.side_effect = FileNotFoundError("Credentials file not found")

        # Act & Assert
        with self.assertRaises(FileNotFoundError) as context:
            connect_to_sheets()
        
        self.assertEqual(str(context.exception), "Credentials file not found")
        mock_credentials.from_service_account_file.assert_called_once()

    @patch('extract_data.service_account.Credentials')
    @patch('extract_data.build')
    def test_build_service_error(self, mock_build, mock_credentials):
        # Arrange
        mock_credentials.from_service_account_file.return_value = MagicMock()
        mock_build.side_effect = Exception("Failed to build service")

        # Act & Assert
        with self.assertRaises(Exception) as context:
            connect_to_sheets()
        
        self.assertEqual(str(context.exception), "Failed to build service")
        mock_credentials.from_service_account_file.assert_called_once()
        mock_build.assert_called_once()

    @patch('extract_data.service_account.Credentials')
    def test_invalid_credentials(self, mock_credentials):
        # Arrange
        mock_credentials.from_service_account_file.side_effect = ValueError("Invalid credentials format")

        # Act & Assert
        with self.assertRaises(ValueError) as context:
            connect_to_sheets()
        
        self.assertEqual(str(context.exception), "Invalid credentials format")
        mock_credentials.from_service_account_file.assert_called_once()

if __name__ == '__main__':
    unittest.main()