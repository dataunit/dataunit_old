import unittest
import model


class TestTestCase(unittest.TestCase):

    def test_init_test_case(self):
        tc = model.TestCase(1,True,"test case name","this is a test")

        self.assertEqual(tc.test_case_id,1)
        self.assertEqual(tc.active, True)
        self.assertEqual(tc.name, "test case name")
        self.assertEqual(tc.description, "this is a test")