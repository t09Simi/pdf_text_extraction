import unittest
import sys

from sparrow_extraction import get_manufacture_model, get_identification_number_list
sys.path.append("..")

class TestSparrowExtraction(unittest.TestCase):
    def test_get_manufacture_model(self):
        # Provide a specific description that matches your data
        description = "Chain Block 6mtr HOL"
        expected_manufacturer, expected_model = "Tiger", "TCB11"
        manufacturer, model = get_manufacture_model(description)
        self.assertEqual(manufacturer, expected_manufacturer)
        self.assertEqual(model, expected_model)

    def test_get_identification_number_list(self):
        # Test case for a specific identification number pattern
        identification_numbers = "S101392"
        expected_output = ["S101392"]
        output = get_identification_number_list(identification_numbers, 12)
        self.assertEqual(output, expected_output)

    # You can add more test methods for other functions or scenarios

if __name__ == '__main__':
    unittest.main()
