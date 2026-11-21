"""Tests for gradefiller/updater.py:  python -m unittest discover tests

No network: downloads read local files through file:// links.
"""
import hashlib
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from gradefiller import updater
from gradefiller.updater import Release


def fake_release(path: Path, **over) -> Release:
    data = path.read_bytes()
    fields = dict(version="9.9.9", url=path.as_uri(), size=len(data), sha256=hashlib.sha256(data).hexdigest())
    fields.update(over)
    return Release(**fields)


class VersionTests(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(updater.parse_version("v1.0.1"), (1, 0, 1))
        self.assertEqual(updater.parse_version("1.10.0"), (1, 10, 0))
        self.assertEqual(updater.parse_version("garbage"), (0,))

    def test_newer(self):
        self.assertTrue(updater.is_newer("1.0.1", "1.0.0"))
        self.assertTrue(updater.is_newer("v1.10.0", "1.9.9"))  # not string order
        self.assertTrue(updater.is_newer("1.0.0", "0.0.0"))  # test builds see releases
        self.assertFalse(updater.is_newer("1.0.0", "1.0.0"))
        self.assertFalse(updater.is_newer("1.0.0", "1.0.1"))


class ParseReleaseTests(unittest.TestCase):
    def test_finds_exe_and_digest(self):
        r = updater.parse_release({
            "tag_name": "v1.0.1",
            "assets": [
                {"name": "notes.txt", "browser_download_url": "x", "size": 1},
                {"name": "GradeFiller.exe", "browser_download_url": "https://e/x.exe",
                 "size": 123, "digest": "sha256:ABCDEF"},
            ],
        })
        self.assertEqual(r, Release("1.0.1", "https://e/x.exe", 123, "abcdef"))

    def test_no_exe(self):
        self.assertIsNone(updater.parse_release({"tag_name": "v2.0.0", "assets": []}))

    def test_no_digest(self):
        r = updater.parse_release({"tag_name": "v1.0.1", "assets": [
            {"name": "GradeFiller.exe", "browser_download_url": "u", "size": 5}]})
        self.assertIsNone(r.sha256)


class DownloadAndSwapTests(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.src = self.dir / "served.exe"
        self.src.write_bytes(b"MZ" + os.urandom(600_000))
        self.dest = self.dir / "GradeFiller.exe.new"

    def tearDown(self):
        for f in self.dir.iterdir():
            f.unlink()
        self.dir.rmdir()

    def test_good_download(self):
        seen = []
        updater.download(fake_release(self.src), self.dest, lambda d, t: seen.append((d, t)))
        self.assertEqual(self.dest.read_bytes(), self.src.read_bytes())
        self.assertEqual(seen[-1], (self.src.stat().st_size,) * 2)
        self.assertFalse(self.dest.with_name(self.dest.name + ".part").exists())

    def test_bad_checksum(self):
        with self.assertRaisesRegex(ValueError, "damaged"):
            updater.download(fake_release(self.src, sha256="0" * 64), self.dest)
        self.assertFalse(self.dest.exists())
        self.assertEqual([p.name for p in self.dir.iterdir()], ["served.exe"])  # no .part left

    def test_short_download(self):
        with self.assertRaisesRegex(ValueError, "stopped early"):
            updater.download(fake_release(self.src, size=10**7, sha256=None), self.dest)

    def test_not_a_program(self):
        self.src.write_bytes(b"<html>rate limited</html>")
        with self.assertRaisesRegex(ValueError, "isn't a Windows program"):
            updater.download(fake_release(self.src), self.dest)

    def test_swap(self):
        exe = self.dir / "GradeFiller.exe"
        exe.write_bytes(b"old")
        self.dest.write_bytes(b"new")
        old = updater.swap_in(self.dest, exe)
        self.assertEqual(exe.read_bytes(), b"new")
        self.assertEqual(old.read_bytes(), b"old")
        self.assertFalse(self.dest.exists())

    def test_swap_puts_old_back_on_failure(self):
        exe = self.dir / "GradeFiller.exe"
        exe.write_bytes(b"old")
        real_replace = os.replace
        calls = []

        def flaky(a, b):
            calls.append((a, b))
            if len(calls) == 2:  # moving the new file in fails
                raise PermissionError("locked by antivirus")
            real_replace(a, b)

        with mock.patch("gradefiller.updater.os.replace", flaky), self.assertRaises(PermissionError):
            updater.swap_in(self.dest, exe)
        self.assertEqual(exe.read_bytes(), b"old")

    def test_install_update(self):
        exe = self.dir / "GradeFiller.exe"
        exe.write_bytes(b"MZold")
        with mock.patch("gradefiller.updater.running_exe", return_value=exe), \
                mock.patch("gradefiller.updater.restart") as restart:
            updater.install_update(fake_release(self.src))
        restart.assert_called_once_with(exe)
        self.assertEqual(exe.read_bytes(), self.src.read_bytes())
        self.assertEqual((self.dir / "GradeFiller.exe.old").read_bytes(), b"MZold")
        self.assertFalse((self.dir / "GradeFiller.exe.new").exists())

    def test_clean_up_old(self):
        exe = self.dir / "GradeFiller.exe"
        exe.write_bytes(b"x")
        (self.dir / "GradeFiller.exe.old").write_bytes(b"x")
        with mock.patch("gradefiller.updater.is_packed_exe", return_value=True), \
                mock.patch("gradefiller.updater.running_exe", return_value=exe):
            updater.clean_up_old()
        self.assertFalse((self.dir / "GradeFiller.exe.old").exists())
        self.assertTrue(exe.exists())


if __name__ == "__main__":
    unittest.main()
