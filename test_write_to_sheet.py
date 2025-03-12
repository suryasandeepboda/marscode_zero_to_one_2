import unittest
from unittest.mock import Mock, patch
import pandas as pd
import numpy as np
from extract_data import write_to_target_sheet

"""Unit tests for Google Sheets write functionality."""

class TestWriteToTargetSheet(unittest.TestCase):
    """Test cases for write_to_target_sheet function."""
    
    def setUp(self):
        # Create sample DataFrame for testing
        self.sample_data = pd.DataFrame({
            'Email Address': ['test1@email.com', 'test2@email.com'],
            'Tool being used': ['Tool1', 'Tool2'],
            'Feature used': ['Feature1', 'Feature2'],
            'Context Awareness': [4.0, 5.0],
            'Autonomy': [3.0, 5.0],
            'Experience': [5.0, 5.0],
            'Output Quality': [4.0, 5.0],
            'Overall Rating': [4.0, 4.0],
            'Unique ID': ['ID1', 'ID2'],
            'Mean Rating': [4.0, 5.0],
            'Difference': [0.0, 1.0],
            'Result': ['Ok', 'Not ok']
        })
        
        # Create mock service
        self.mock_service = Mock()
        self.mock_sheets = Mock()
        self.mock_service.spreadsheets.return_value = self.mock_sheets

    def test_successful_write(self):
        # Configure mocks
        self.mock_sheets.values().clear().execute.return_value = {}
        self.mock_sheets.values().update().execute.return_value = {}
        self.mock_sheets.batchUpdate().execute.return_value = {}
        
        # Execute function
        result = write_to_target_sheet(self.mock_service, self.sample_data)
        
        # Verify result
        self.assertTrue(result)
        
        # Verify clear was called
        self.mock_sheets.values().clear.assert_called_once()
        
        # Verify update was called
        self.mock_sheets.values().update.assert_called_once()
        
        # Verify batch update for formatting was called
        self.mock_sheets.batchUpdate.assert_called_once()

    def test_handle_empty_values(self):
        # Create DataFrame with empty values
        data_with_empty = self.sample_data.copy()
        data_with_empty.loc[0, 'Context Awareness'] = np.nan
        
        # Configure mocks
        self.mock_sheets.values().clear().execute.return_value = {}
        self.mock_sheets.values().update().execute.return_value = {}
        self.mock_sheets.batchUpdate().execute.return_value = {}
        
        # Execute function
        result = write_to_target_sheet(self.mock_service, data_with_empty)
        
        # Verify result
        self.assertTrue(result)

    def test_handle_clear_failure(self):
        # Configure mock to fail on clear
        self.mock_sheets.values().clear().execute.side_effect = Exception("Clear failed")
        
        # Execute function
        result = write_to_target_sheet(self.mock_service, self.sample_data)
        
        # Verify result
        self.assertFalse(result)
        
        # Verify clear was attempted
        self.mock_sheets.values().clear.assert_called_once()
        
        # Verify update was not called
        self.mock_sheets.values().update.assert_not_called()

    def test_handle_update_failure(self):
        # Configure mocks
        self.mock_sheets.values().clear().execute.return_value = {}
        self.mock_sheets.values().update().execute.side_effect = Exception("Update failed")
        
        # Execute function
        result = write_to_target_sheet(self.mock_service, self.sample_data)
        
        # Verify result
        self.assertFalse(result)
        
        # Verify clear was called
        self.mock_sheets.values().clear.assert_called_once()
        
        # Verify update was attempted
        self.mock_sheets.values().update.assert_called_once()

    def test_handle_formatting_failure(self):
        # Configure mocks
        self.mock_sheets.values().clear().execute.return_value = {}
        self.mock_sheets.values().update().execute.return_value = {}
        self.mock_sheets.batchUpdate().execute.side_effect = Exception("Formatting failed")
        
        # Execute function
        result = write_to_target_sheet(self.mock_service, self.sample_data)
        
        # Verify result
        self.assertFalse(result)
        
        # Verify previous operations were called
        self.mock_sheets.values().clear.assert_called_once()
        self.mock_sheets.values().update.assert_called_once()
        self.mock_sheets.batchUpdate.assert_called_once()

if __name__ == '__main__':
    unittest.main()