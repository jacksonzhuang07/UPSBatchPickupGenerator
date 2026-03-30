import unittest
from address_parser import parse_address_string, split_addresses

class TestAddressParser(unittest.TestCase):
    def test_standard_address(self):
        addr = "123 Main St, Springfield, IL 62704"
        parsed = parse_address_string(addr)
        self.assertEqual(parsed['Street'], "123 Main St")
        self.assertEqual(parsed['City'], "Springfield")
        self.assertEqual(parsed['State'], "IL")
        self.assertEqual(parsed['Zip'], "62704")

    def test_address_with_suite(self):
        addr = "456 Oak Ave Suite 10, Los Angeles, CA 90001"
        parsed = parse_address_string(addr)
        self.assertEqual(parsed['Street'], "456 Oak Ave")
        self.assertEqual(parsed['Suite'], "Suite 10")
        self.assertEqual(parsed['City'], "Los Angeles")
        self.assertEqual(parsed['State'], "CA")
        self.assertEqual(parsed['Zip'], "90001")

    def test_canadian_address(self):
        addr = "2320 HWY. NO. 2, BOWMANVILLE, ON L1C3K5, Canada"
        parsed = parse_address_string(addr)
        self.assertEqual(parsed['Country'], "CA")
        self.assertEqual(parsed['State'], "ON")
        self.assertEqual(parsed['Zip'], "L1C3K5")
        self.assertTrue("2320" in parsed['Street'])

    def test_split_single_newlines(self):
        batch = """60 PINE GROVE RD, BRIDGEWATER, NS B4V4H2, Canada
1400 Ottawa st S, Kitchener, ON N2E4E2, Canada
3355 JOHNSTON ROAD (HWY. 4), PORT ALBERNI, BC V9Y8E9, Canada"""
        split = split_addresses(batch)
        self.assertEqual(len(split), 3)

    def test_split_mixed_newlines(self):
        batch = """2320 HWY. NO. 2, BOWMANVILLE, ON L1C3K5, Canada

60 PINE GROVE RD, BRIDGEWATER, NS B4V4H2, Canada
1400 Ottawa st S, Kitchener, ON N2E4E2, Canada"""
        split = split_addresses(batch)
        self.assertEqual(len(split), 3)

    def test_edson_address(self):
        addr = "5750-2ND AVENUE, EDSON, AB T7E1M2, Canada"
        parsed = parse_address_string(addr)
        self.assertEqual(parsed['Country'], "CA")
        self.assertEqual(parsed['State'], "AB")
        # Check if either full street or at least the number part is there
        self.assertTrue(len(parsed['Street']) > 0, "Street should not be empty")
        self.assertTrue("5750" in parsed['Street'])

if __name__ == "__main__":
    unittest.main()
