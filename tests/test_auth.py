import unittest

from app import APP_PASSWORD, auth_cookie_value, is_auth_cookie_valid, login_page


class AuthTests(unittest.TestCase):
    def test_default_password_is_judy(self):
        self.assertEqual(APP_PASSWORD, "judy")

    def test_auth_cookie_value_must_match_signature(self):
        cookie = auth_cookie_value()

        self.assertTrue(is_auth_cookie_valid(cookie))
        self.assertFalse(is_auth_cookie_valid(""))
        self.assertFalse(is_auth_cookie_valid("v1:not-the-signature"))

    def test_login_page_can_render_error(self):
        rendered = login_page("Incorrect password.")

        self.assertIn("Shiftline", rendered)
        self.assertIn("Incorrect password.", rendered)
        self.assertIn('name="password"', rendered)


if __name__ == "__main__":
    unittest.main()
