import unittest
from src.first_integrated import split_id_numbers_with_range, add_six_months, get_manufacture_model, \
    extract_first_integrated_pdf


class TestFunctions(unittest.TestCase):
    def test_split_id_numbers_with_range(self):
        id_numbers = ['TZW238-243', 'TZW227-231']
        expected_result = (['TZW238', 'TZW239', 'TZW240', 'TZW241', 'TZW242', 'TZW243', 'TZW227', 'TZW228', 'TZW229',
                            'TZW230', 'TZW231'], [])
        self.assertEqual(split_id_numbers_with_range(id_numbers), expected_result)

    def test_add_six_months(self):
        date_str = '23/02/2024'
        expected_result = '21/08/2024'  # Adding 6 months to January 1, 2023
        self.assertEqual(add_six_months(date_str), expected_result)

    def test_get_manufacture_model(self):
        # Prepare test data
        description = 'Wire Rope Pulling Machine'
        expected_manufacturer, expected_model = "", ""
        manufacturer, model = get_manufacture_model(description)
        self.assertEqual(manufacturer, expected_manufacturer)
        self.assertEqual(model, expected_model)

    def test_extract_first_integrated_pdf(self):
        pdf_path = "../../resources/First Integrated Full Cert Pack.pdf"
        result = extract_first_integrated_pdf(pdf_path)
        self.assertEqual(None, result)


if __name__ == '__main__':
    unittest.main()
