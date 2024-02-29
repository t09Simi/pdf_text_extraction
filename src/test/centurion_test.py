import unittest
import sys

from ..centurion_extraction import get_manufacture, extract_quantity
sys.path.append("..")


class TestFunctions(unittest.TestCase):
    def test_get_manufacture(self):
        # Assume there is a valid description and expected output
        description = "CHAINBLOCK 1T 3M TIGER TCB14"
        expected_manufacturer = "TIGER"
        expected_model = "TCB14"
        manufacturer, model = get_manufacture(description)
        self.assertEqual(manufacturer, expected_manufacturer)
        self.assertEqual((model, expected_model))

    def test_extract_quantity(self):
        text = " 4 RIDGEGEAR RGL28 2MTR TWIN TAIL LANYARD"
        expected_quantity = 4
        quantity = extract_quantity(text)
        self.assertEqual(quantity, expected_quantity)


if __name__ == '__main__':
    unittest.main()




