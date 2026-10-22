"""The dev lock gate passes the committed lock and fails one compiled for
Linux alone, as Dependabot's pip-compile writes it."""

from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
import check_dev_lock  # noqa: E402

# How Dependabot's pip-compile writes the lock: no markers, and no package
# only Windows needs.
LINUX_ONLY = """\
cffi==2.1.1 \\
    --hash=sha256:046bfc24911b37851ee1b51aab8bffe713d89c68c6a057b09484ce9fd5f69b4e
    # via cryptography
pcodedmp==1.2.6 \\
    --hash=sha256:025f8c809a126f45a082ffa820893e6a8d990d9d7ddb68694b5a9f0a6dbcd955
    # via oletools
"""


class DevLockTests(unittest.TestCase):
    def test_the_committed_lock_pins_its_windows_packages(self):
        text = check_dev_lock.LOCK.read_text(encoding="utf-8")
        self.assertEqual(check_dev_lock.missing_windows_packages(text), [])

    def test_a_lock_compiled_for_linux_alone_fails(self):
        self.assertEqual(check_dev_lock.missing_windows_packages(LINUX_ONLY),
                         list(check_dev_lock.WINDOWS_PACKAGES))

    def test_a_windows_package_pinned_without_its_marker_fails(self):
        text = check_dev_lock.LOCK.read_text(encoding="utf-8").replace(
            "pywin32==312 ; sys_platform == 'win32'", "pywin32==312")
        self.assertEqual(check_dev_lock.missing_windows_packages(text), ["pywin32"])

    def test_names_compare_as_pip_normalizes_them(self):
        text = "Win_Unicode_Console==0.5 ; sys_platform == 'win32' \\\n"
        self.assertNotIn("win-unicode-console",
                         check_dev_lock.missing_windows_packages(text))


if __name__ == "__main__":
    unittest.main()
