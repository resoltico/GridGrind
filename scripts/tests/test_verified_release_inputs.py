"""Prior release evidence must fail closed before any checks can be reused."""
import copy
import importlib.util
from pathlib import Path
import unittest
import urllib.request

SPEC = importlib.util.spec_from_file_location("verified_release_inputs", Path(__file__).parents[1] / "verified-release-inputs.py")
resume = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(resume)
COMMIT = "1" * 40
AMD64 = "sha256:" + "2" * 64
ARM64 = "sha256:" + "3" * 64


class Api:
    def __init__(self):
        self.run = {"id": 42, "run_attempt": 1, "repository": {"full_name": "owner/GridGrind"},
                    "path": ".github/workflows/release.yml", "head_sha": COMMIT,
                    "head_branch": "v0.77.0", "event": "push", "status": "completed", "conclusion": "failure"}
        self.job = {"id": 7, "run_id": 42, "run_attempt": 1, "name": "release", "status": "completed", "conclusion": "failure",
                    "steps": [{"name": name, "status": "completed", "conclusion": "success"} for name in resume.REQUIRED]
                             + [{"name": resume.PUBLICATION, "status": "completed", "conclusion": "failure"}]}
        command = './gradlew containerTest "-Pversion=0.77.0" "-PcontainerImage=%s"'
        self.log = ("2026-10-28T00:00:00Z ##[group]Run " + command % AMD64 + "\n"
                    + "2026-10-28T00:00:00Z \x1b[36;1m" + command % AMD64 + "\x1b[0m\n"
                    + "2026-10-28T00:00:00Z " + command % ARM64 + "\n"
                    + "2026-10-28T00:00:00Z ##[endgroup]\n")

    def get(self, route, log=False):
        if log:
            return self.log
        if route.endswith("/jobs?per_page=100"):
            return {"total_count": 1, "jobs": [copy.deepcopy(self.job)]}
        return copy.deepcopy(self.run)


class VerifiedInputsTest(unittest.TestCase):
    def verify(self, api):
        return resume.verified_images(api, "owner/GridGrind", "42", "0.77.0", COMMIT)

    def test_actual_command_header_duplication_does_not_confuse_platform_order(self):
        self.assertEqual({"amd64": AMD64, "arm64": ARM64}, self.verify(Api()))

    def test_wrong_commit_workflow_repository_event_or_attempt_is_rejected(self):
        for key, value in (("head_sha", "4" * 40), ("path", ".github/workflows/ci.yml"),
                           ("repository", {"full_name": "fork/GridGrind"}), ("event", "pull_request"),
                           ("head_branch", "main"), ("run_attempt", 2), ("status", "in_progress")):
            api = Api()
            api.run[key] = value
            with self.subTest(key=key), self.assertRaises(ValueError):
                self.verify(api)

    def test_every_required_step_must_have_passed(self):
        for index in range(len(resume.REQUIRED)):
            for conclusion in ("failure", "skipped", "cancelled", None):
                api = Api()
                api.job["steps"][index]["conclusion"] = conclusion
                with self.subTest(index=index, conclusion=conclusion), self.assertRaises(ValueError):
                    self.verify(api)

    def test_failure_elsewhere_or_duplicate_step_is_rejected(self):
        api = Api()
        api.job["steps"][-1]["name"] = "Publish runtime"
        with self.assertRaises(ValueError):
            self.verify(api)
        api = Api()
        api.job["steps"].append(copy.deepcopy(api.job["steps"][0]))
        with self.assertRaises(ValueError):
            self.verify(api)

    def test_missing_duplicate_or_wrong_version_image_commands_are_rejected(self):
        for change in (lambda log: log.replace(ARM64, AMD64),
                       lambda log: log.replace('"-Pversion=0.77.0"', '"-Pversion=0.76.0"'),
                       lambda log: log + log,
                       lambda log: log.replace("##[group]Run", "unrelated")):
            api = Api()
            api.log = change(api.log)
            with self.assertRaises(ValueError):
                self.verify(api)

    def test_unrelated_output_cannot_supply_missing_verification_images(self):
        api = Api()
        api.log += '2026-10-28T00:00:00Z ./gradlew containerTest "-Pversion=0.77.0" "-PcontainerImage=sha256:' + "4" * 64 + '"\n'
        self.assertEqual({"amd64": AMD64, "arm64": ARM64}, self.verify(api))
        api.log = api.log.replace('"-PcontainerImage=' + ARM64 + '"', '"missing"')
        with self.assertRaises(ValueError):
            self.verify(api)

    def test_log_redirect_strips_credentials_and_rejects_insecure_targets(self):
        request = urllib.request.Request("https://api.github.com/logs", headers={
            "Authorization": "Bearer fixture-token", "Cookie": "fixture-cookie"})
        redirect = resume.LogRedirect()
        target = redirect.redirect_request(request, None, 302, "Found", {}, "https://logs.example.test/log")
        self.assertIsNone(target.get_header("Authorization"))
        self.assertIsNone(target.get_header("Cookie"))
        with self.assertRaises(ValueError):
            redirect.redirect_request(request, None, 302, "Found", {}, "http://logs.example.test/log")
        with self.assertRaises(ValueError):
            resume.NoRedirect().redirect_request(request, None, 302, "Found", {}, "https://logs.example.test/log")


if __name__ == "__main__":
    unittest.main()
