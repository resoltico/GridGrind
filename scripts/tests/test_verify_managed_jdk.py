"""The managed JDK gate rejects wrong cached or mislabeled binary identities."""

import importlib.util
from pathlib import Path
import unittest

MODULE_PATH = Path(__file__).resolve().parents[1] / "verify-managed-jdk.py"
SPEC = importlib.util.spec_from_file_location("verify_managed_jdk", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class ManagedJdkIdentityTest(unittest.TestCase):
    def setUp(self):
        self.pins = {
            "TEMURIN_VERSION": "27.0.1.0.2+1",
            "TEMURIN_SOURCE_COMMIT": "a" * 40,
            "TEMURIN_BUILD_COMMIT": "b" * 40,
        }
        self.release = {
            "IMPLEMENTOR": "Eclipse Adoptium",
            "IMPLEMENTOR_VERSION": "Temurin-27.0.1.0.2+1",
            "JAVA_RUNTIME_VERSION": "27.0.1.0.2+1",
            "SOURCE": ".:git:" + "a" * 12,
            "BUILD_SOURCE": "git:" + "b" * 40,
        }
        self.runtime = {
            "java.runtime.version": "27.0.1.0.2+1",
            "java.vendor": "Eclipse Adoptium",
        }

    def test_accepts_the_exact_runtime_and_matching_source_build_tuple(self):
        MODULE.verify_identity(self.pins, self.release, self.runtime)

    def test_rejects_wrong_release_vendor_version_and_build_source(self):
        for field in ("IMPLEMENTOR", "IMPLEMENTOR_VERSION", "JAVA_RUNTIME_VERSION", "BUILD_SOURCE"):
            with self.subTest(field=field):
                release = dict(self.release, **{field: "wrong"})
                with self.assertRaisesRegex(ValueError, field):
                    MODULE.verify_identity(self.pins, release, self.runtime)

    def test_rejects_foreign_or_missing_matching_source_commit(self):
        for source in (".:git:" + "c" * 12, ".:git:a", ""):
            with self.subTest(source=source):
                with self.assertRaisesRegex(ValueError, "SOURCE"):
                    MODULE.verify_identity(self.pins, dict(self.release, SOURCE=source), self.runtime)

    def test_rejects_wrong_running_vendor_or_abbreviated_runtime_build(self):
        for key, value in (("java.vendor", "Other vendor"), ("java.runtime.version", "27.0.1+0")):
            with self.subTest(key=key):
                with self.assertRaises(ValueError):
                    MODULE.verify_identity(self.pins, self.release, dict(self.runtime, **{key: value}))

    def test_missing_pins_fail_with_an_actionable_error(self):
        with self.assertRaisesRegex(ValueError, "Missing managed JDK pins"):
            MODULE.verify_identity({}, self.release, self.runtime)


if __name__ == "__main__":
    unittest.main()
