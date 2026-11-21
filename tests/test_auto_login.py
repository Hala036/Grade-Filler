"""Auto-login against a fake portal + fake university login page, in a hidden browser.

    python -m unittest discover tests

The fake portal lives on "localhost" and the fake login page on "127.0.0.1"
(different hosts, like portal.alquds.edu and login.alquds.edu). Skipped when no
browser can be started.
"""
import tempfile
import threading
import time
import unittest
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from gradefiller.automation import PortalSession
from gradefiller.credentials import Login

GOOD = ("teacher1", "s3cret")


class FakeSites(BaseHTTPRequestHandler):
    posts = 0

    def log_message(self, *a):
        pass

    def _send(self, code=200, body="", headers=()):
        self.send_response(code)
        for k, v in headers:
            self.send_header(k, v)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(body.encode())

    def _form(self, username="", error=""):
        return (f"<html><body>{error}<form action='/sso' method='post' name='f'>"
                f"<input id='username' name='username' value='{username}'>"
                "<input id='password' type='password' name='password'>"
                "<button id='submit_button' class='btn'>Login</button></form></body></html>")

    def do_GET(self):
        host = self.headers["Host"].split(":")[0]
        port = self.server.server_port
        if host == "localhost" and self.path == "/course_list":
            if "sess=ok" in (self.headers.get("Cookie") or ""):
                return self._send(body="<h1>COURSES</h1>")
            return self._send(302, headers=[("Location", f"http://127.0.0.1:{port}/sso")])
        if host == "localhost" and self.path == "/welcome":  # where the portal lands after login
            return self._send(body="<h1>WELCOME</h1>")
        if host == "localhost" and self.path == "/set":
            return self._send(302, headers=[("Set-Cookie", "sess=ok; Path=/"), ("Location", "/welcome")])
        if host == "127.0.0.1" and self.path == "/sso":
            return self._send(body=self._form())
        self._send(404)

    def do_POST(self):
        FakeSites.posts += 1
        data = urllib.parse.parse_qs(self.rfile.read(int(self.headers["Content-Length"])).decode())
        user, pw = data.get("username", [""])[0], data.get("password", [""])[0]
        if (user, pw) == GOOD:
            return self._send(302, headers=[("Location", f"http://localhost:{self.server.server_port}/set")])
        self._send(body=self._form(user, "<div class='error'>Incorrect username or password</div>"))


class AutoLoginTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), FakeSites)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        cls.course = f"http://localhost:{cls.server.server_port}/course_list"

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        FakeSites.posts = 0
        self.sessions = []

    def tearDown(self):
        for s in self.sessions:
            s.close()

    def session(self, saved=None) -> PortalSession:
        s = PortalSession(self.tmp / "state.json", saved_login=saved, sso_host="127.0.0.1")
        s.HEADLESS = True
        self.sessions.append(s)
        try:
            s._ensure_browser()
        except RuntimeError as e:
            self.skipTest(str(e))
        return s

    def test_typed_login_is_noticed(self):
        s = self.session()
        s.open_course_list(self.course)
        self.assertTrue(s.on_login_page())
        # The teacher types and presses Log in while Python is idle (as it is
        # while the app waits for them); Python only looks again afterwards.
        s.page.evaluate("""() => setTimeout(() => {
            document.querySelector('#username').value = 'teacher1';
            document.querySelector('#password').value = 's3cret';
            document.querySelector('#submit_button').click();
        }, 300)""")
        time.sleep(2)
        self.assertIn("WELCOME", s.page.content())
        self.assertEqual(s.typed_login, GOOD)

    def test_saved_login_signs_in(self):
        s = self.session(Login(*GOOD))
        s.open_course_list(self.course)
        self.assertIn("COURSES", s.page.content())
        self.assertEqual(FakeSites.posts, 1)
        self.assertEqual(s.typed_login, GOOD)  # same as saved: the app won't ask again

    def test_wrong_saved_login_tries_once(self):
        s = self.session(Login("teacher1", "old-password"))
        started = time.monotonic()
        s.open_course_list(self.course)
        self.assertLess(time.monotonic() - started, 8)  # gives up quickly, not after 20 s
        self.assertTrue(s.on_login_page())
        self.assertIn("Incorrect", s.page.content())
        s.open_course_list(self.course)  # back to the login page within 2 minutes
        time.sleep(1)
        self.assertTrue(s.on_login_page())
        self.assertEqual(s.page.input_value("#password"), "old-password")  # filled in, not pressed
        self.assertEqual(FakeSites.posts, 1)  # never pressed again

    def test_switch_account_logs_out(self):
        s = self.session(Login(*GOOD))
        s.open_course_list(self.course)
        self.assertIn("COURSES", s.page.content())
        s.set_saved_login(None)
        s.log_out()
        s.open_course_list(self.course)
        self.assertTrue(s.on_login_page())
        self.assertEqual(s.page.input_value("#password"), "")  # nothing filled in


if __name__ == "__main__":
    unittest.main()
