import unittest
import sys

from centurion_extraction import get_manufacture, extract_quantity, extraction_centurion_pdf, \
    get_identification_parts_list, get_identification_number_list

sys.path.append("..")


class TestFunctions(unittest.TestCase):
    def test_get_manufacture(self):
        # Assume there is a valid description and expected output
        description = "CHAINBLOCK 1T 3M TIGER TCB14"
        expected_manufacturer = "TIGER"
        expected_model = "TCB14"
        manufacturer, model = get_manufacture(description)
        self.assertEqual(manufacturer.lower(), expected_manufacturer.lower())
        self.assertEqual(model.lower(), expected_model.lower())

    def test_extract_quantity(self):
        text = " 4 RIDGEGEAR RGL28 2MTR TWIN TAIL LANYARD"
        expected_quantity = 4
        quantity = extract_quantity(text)
        self.assertEqual(quantity, expected_quantity)

    def test_extraction_centurion_pdf(self):
        pdf_path = "../resources/centurion.pdf"
        result = extraction_centurion_pdf(pdf_path)
        self.assertEqual(None, result)

    def test_get_identification_parts_list(self):
        result = get_identification_parts_list("1", 4)
        self.assertEqual(["1", "2", "3", "4"], result)

    def test_get_identification_number_list(self):
        result = get_identification_number_list('D971-1 to MGL1', 2)
        self.assertEqual(['D971-1', 'D971-2'], result)


if __name__ == '__main__':
    unittest.main()