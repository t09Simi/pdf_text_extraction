import unittest
from src.first_integrated import split_id_numbers_with_range, add_six_months


class TestFunctions(unittest.TestCase):
    def test_split_id_numbers_with_range(self):
        id_numbers = ['TZW238-243', 'TZW227-231']
        expected_result = ['TZW238', 'TZW239', 'TZW240', 'TZW241', 'TZW242', 'TZW243', 'TZW227', 'TZW228', 'TZW229',
                           'TZW230', 'TZW231']
        self.assertEqual(split_id_numbers_with_range(id_numbers), expected_result)

    def test_add_six_months(self):
        date_str = '23/02/2024'
        expected_result = '21/08/2024'  # Adding 6 months to January 1, 2023
        self.assertEqual(add_six_months(date_str), expected_result)


if __name__ == '__main__':
    unittest.main()
