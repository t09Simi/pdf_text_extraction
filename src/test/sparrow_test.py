import unittest
import sys
import os

current_directory = os.getcwd()

sys.path.append(os.path.join(current_directory,'src'))
from sparrow_extraction import get_manufacture_model, get_identification_number_list,get_identification_parts_list,extract_sparrow_pdf



class TestSparrowExtraction(unittest.TestCase):
    def test_get_manufacture_model(self):
        # Provide a specific description that matches your data
        description = "Chain Block 6mtr HOL"
        expected_manufacturer, expected_model = "", ""
        manufacturer, model = get_manufacture_model(description)
        self.assertEqual(manufacturer, expected_manufacturer)
        self.assertEqual(model, expected_model)

    def test_get_identification_number_list(self):
        # Test case for a specific identification number pattern
        identification_numbers = "D971-1 to MGL1"
        expected_output = ["D971-1", "D971-2", "D971-3"]
        output = get_identification_number_list(identification_numbers, 3)
        self.assertEqual(output, expected_output)
    def test_get_identification_parts_list(self):
        result = get_identification_parts_list("1", 4)
        self.assertEqual(["1", "2", "3", "4"], result)

    def test_extract_sparrow_pdf(self):
        pdf_path = "resources/centurion.pdf"
        result = extract_sparrow_pdf(pdf_path)
        self.assertEqual(None, result)

    # You can add more test methods for other functions or scenarios


if __name__ == '__main__':
    unittest.main()
