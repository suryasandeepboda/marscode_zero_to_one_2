import unittest
from unittest.mock import Mock, patch
import pandas as pd
import numpy as np
from extract_data import extract_sheet_data

class TestExtractSheetData(unittest.TestCase):
    def setUp(self):
        self.sample_data = [
            ['Email Address', 'Tool being used', 'Feature used', 'Context Awareness', 'Autonomy', 'Experience', 'Output Quality', 'Overall Rating', 'Unique ID'],
            ['test@email.com', 'Tool1', 'Feature1', '4', '3', '5', '4', '4', 'ID1'],
            ['test2@email.com', 'Tool2', 'Feature2', '5', '5', '5', '5', '4', 'ID2']
        ]

    @patch('extract_data.connect_to_sheets')
    def test_successful_data_extraction(self, mock_connect):
        # Mock the Google Sheets service and response
        mock_service = Mock()
        mock_sheet = Mock()
        mock_values = Mock()
        
        mock_connect.return_value = mock_service
        mock_service.spreadsheets.return_value = mock_sheet
        mock_sheet.values.return_value.get.return_value.execute.return_value = {'values': self.sample_data}
        
        # Execute function
        result = extract_sheet_data()
        
        # Assertions
        self.assertIsNotNone(result)
        self.assertIsInstance(result, pd.DataFrame)
        self.assertEqual(len(result), 2)  # Two data rows
        self.assertTrue('Mean Rating' in result.columns)
        self.assertTrue('Difference' in result.columns)
        self.assertTrue('Result' in result.columns)

    @patch('extract_data.connect_to_sheets')
    def test_empty_sheet(self, mock_connect):
        # Mock empty response
        mock_service = Mock()
        mock_sheet = Mock()
        mock_values = Mock()
        
        mock_connect.return_value = mock_service
        mock_service.spreadsheets.return_value = mock_sheet
        mock_sheet.values.return_value.get.return_value.execute.return_value = {'values': []}
        
        # Execute function
        result = extract_sheet_data()
        
        # Assertions
        self.assertIsNone(result)

    @patch('extract_data.connect_to_sheets')
    def test_missing_required_columns(self, mock_connect):
        # Data with missing columns
        incomplete_data = [
            ['Email Address', 'Tool being used', 'Feature used'],  # Missing required columns
            ['test@email.com', 'Tool1', 'Feature1']
        ]
        
        mock_service = Mock()
        mock_sheet = Mock()
        mock_values = Mock()
        
        mock_connect.return_value = mock_service
        mock_service.spreadsheets.return_value = mock_sheet
        mock_sheet.values.return_value.get.return_value.execute.return_value = {'values': incomplete_data}
        
        # Execute function and expect ValueError
        with self.assertRaises(ValueError):
            result = extract_sheet_data()

    @patch('extract_data.connect_to_sheets')
    def test_result_calculation(self, mock_connect):
        # Test data with known results
        test_data = [
            ['Email Address', 'Tool being used', 'Feature used', 'Context Awareness', 'Autonomy', 'Experience', 'Output Quality', 'Overall Rating', 'Unique ID'],
            ['test@email.com', 'Tool1', 'Feature1', '5', '5', '5', '5', '4', 'ID1'],  # Mean=5, Diff=1, Result=Ok
            ['test2@email.com', 'Tool2', 'Feature2', '5', '5', '5', '5', '2', 'ID2']  # Mean=5, Diff=3, Result=Not ok
        ]
        
        mock_service = Mock()
        mock_sheet = Mock()
        mock_values = Mock()
        
        mock_connect.return_value = mock_service
        mock_service.spreadsheets.return_value = mock_sheet
        mock_sheet.values.return_value.get.return_value.execute.return_value = {'values': test_data}
        
        # Execute function
        result = extract_sheet_data()
        
        # Assertions
        self.assertIsNotNone(result)
        self.assertEqual(result.iloc[0]['Result'], 'Ok')
        self.assertEqual(result.iloc[1]['Result'], 'Not ok')
        self.assertEqual(result.iloc[0]['Mean Rating'], 5.0)
        self.assertEqual(result.iloc[0]['Difference'], 1.0)

    @patch('extract_data.connect_to_sheets')
    def test_api_error_handling(self, mock_connect):
        # Mock API error
        mock_connect.side_effect = Exception("API Error")
        
        # Execute function
        result = extract_sheet_data()
        
        # Assertions
        self.assertIsNone(result)

if __name__ == '__main__':
    unittest.main()