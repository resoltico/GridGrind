"""Pure publication-state fixtures: never contact GitHub or upload files."""

import copy
import importlib.util
import io
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock
import urllib.error

SPEC = importlib.util.spec_from_file_location("release_publication", Path(__file__).parents[1] / "release-publication.py")
publication = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = publication
SPEC.loader.exec_module(publication)
COMMIT = "1" * 40
NOTES = "Reviewed release summary.\n"
DIGEST = "sha256:ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
ASSET = publication.Asset(Path("/fixture/gridgrind.jar"), "gridgrind.jar", 3, DIGEST)


def release(draft=True, prerelease=True, body=NOTES):
    return {"id": 7, "tag_name": "v0.77.0", "draft": draft, "prerelease": prerelease, "body": body}


def remote_asset():
    return {"name": ASSET.name, "size": ASSET.size, "digest": ASSET.digest, "state": "uploaded"}


class Api:
    def __init__(self, current=None, assets=None):
        self.current = current
        self.assets = [remote_asset()] if assets is None else assets
        self.mutations = []
        self.sha = COMMIT
        self.lookup_error = None
        self.annotated = False
        self.list_pages = None
        self.list_error = None
        self.reads = []

    def request(self, method, route, body=None):
        if method == "GET":
            self.reads.append(route)
            if route.startswith("/git/ref/tags/"):
                return {"object": {"type": "tag" if self.annotated else "commit", "sha": self.sha}}
            if route.startswith("/git/tags/"):
                return {"object": {"type": "commit", "sha": COMMIT}}
            if route.startswith("/releases/tags/"):
                if self.lookup_error:
                    raise publication.ApiError(self.lookup_error)
                if self.current is None or self.current["draft"]:
                    raise publication.ApiError(404)
                return copy.deepcopy(self.current)
            if route.startswith("/releases?per_page=100&page="):
                if self.list_error:
                    raise publication.ApiError(self.list_error)
                page = int(route.rsplit("=", 1)[1])
                if self.list_pages is not None:
                    return copy.deepcopy(self.list_pages.get(page, []))
                return [] if self.current is None else [copy.deepcopy(self.current)]
            if self.current is not None and route.startswith("/releases/%s/assets?" % self.current["id"]):
                return copy.deepcopy(self.assets)
            if route == "/releases/latest":
                return copy.deepcopy(self.current)
        self.mutations.append((method, route, body))
        if method == "POST" and route == "/releases":
            self.current = {"id": 7, **body}
        elif method == "PATCH" and self.current is not None and route == "/releases/" + str(self.current["id"]):
            self.current.update(body)
            # Observed GitHub draft behavior: an omitted tag resets its identity.
            self.current["tag_name"] = body.get("tag_name", "untagged-fixture")
        else:
            raise AssertionError("Unexpected fixture request: " + method + " " + route)
        return copy.deepcopy(self.current)


class PublicationTest(unittest.TestCase):
    def helper(self, api):
        self.uploads = []
        def upload(repo, tag, assets):
            self.assertTrue(api.current["draft"], "Only draft assets may be overwritten")
            self.uploads.append((repo, tag))
            api.assets = [remote_asset()]
        return publication.Publication(api, "owner/GridGrind", "0.77.0", COMMIT, NOTES, (ASSET,), upload)

    def test_prepare_creates_verified_draft(self):
        api = Api()
        self.helper(api).run("prepare")
        self.assertTrue(api.current["draft"])
        self.assertTrue(api.current["prerelease"])
        self.assertEqual("false", api.current["make_latest"])
        self.assertEqual([("owner/GridGrind", "v0.77.0")], self.uploads)

    def test_prepare_resumes_draft_and_replaces_only_draft_assets(self):
        api = Api(release(body="outdated"), [{"name": ASSET.name}])
        self.helper(api).run("prepare")
        self.assertEqual(NOTES, api.current["body"])
        self.assertTrue(api.current["draft"])
        self.assertEqual(1, len(self.uploads))

    def test_tag_404_finds_existing_listed_draft_and_does_not_create_again(self):
        current = release(body="")
        current["id"] = 426403460
        api = Api(current, [])
        self.helper(api).run("prepare")
        self.assertEqual(426403460, api.current["id"])
        self.assertTrue(api.current["draft"])
        self.assertEqual(NOTES, api.current["body"])
        self.assertIn("/releases?per_page=100&page=1", api.reads)
        self.assertFalse(any(method == "POST" for method, _, _ in api.mutations))
        self.assertEqual(1, len(self.uploads))

    def test_duplicate_matching_drafts_across_pages_abort_before_mutation(self):
        api = Api()
        first = release()
        second = {**release(), "id": 8}
        unrelated = [{**release(), "id": number, "tag_name": "unrelated-" + str(number)} for number in range(10, 109)]
        api.list_pages = {1: unrelated + [first], 2: [second]}
        with self.assertRaisesRegex(publication.PublicationError, "Multiple GitHub releases"):
            self.helper(api).run("prepare")
        self.assertEqual([], api.mutations)
        self.assertEqual([], self.uploads)

    def test_duplicate_paginated_release_id_aborts(self):
        api = Api()
        page = [{**release(), "id": number, "tag_name": "unrelated-" + str(number)} for number in range(10, 110)]
        api.list_pages = {1: page, 2: [page[0]]}
        with self.assertRaisesRegex(publication.PublicationError, "duplicate IDs"):
            self.helper(api).run("prepare")
        self.assertEqual([], api.mutations)

    def test_invalid_list_pages_or_release_identity_abort(self):
        for invalid in ({"message": "not a page"}, [release()] * 101, [{"id": 7}],
                        [{**release(), "id": True}], [{**release(), "draft": "true"}]):
            api = Api()
            api.list_pages = {1: invalid}
            with self.assertRaises(publication.PublicationError):
                self.helper(api).run("prepare")
            self.assertEqual([], api.mutations)

    def test_release_list_errors_do_not_prove_absence(self):
        for status in (403, 404, 500):
            api = Api()
            api.list_error = status
            with self.assertRaises(publication.ApiError):
                self.helper(api).run("prepare")
            self.assertEqual([], api.mutations)

    def test_complete_absent_list_allows_creation_and_then_finds_new_draft(self):
        api = Api()
        self.helper(api).run("prepare")
        self.assertEqual(1, sum(method == "POST" for method, _, _ in api.mutations))
        self.assertTrue(api.current["draft"])
        self.assertGreaterEqual(api.reads.count("/releases?per_page=100&page=1"), 2)

    def test_every_patch_sends_complete_explicit_identity_and_desired_state(self):
        api = Api()
        helper = self.helper(api)
        for stage in ("prepare", "expose", "finalize"):
            helper.run(stage)
        patches = [body for method, _, body in api.mutations if method == "PATCH"]
        self.assertEqual(3, len(patches))
        for body, state in zip(patches, ((True, True, "false"), (False, True, "false"), (False, False, "true"))):
            self.assertEqual({
                "tag_name": "v0.77.0", "target_commitish": COMMIT,
                "name": "GridGrind 0.77.0", "body": NOTES,
                "draft": state[0], "prerelease": state[1], "make_latest": state[2],
            }, body)
        self.assertEqual("v0.77.0", api.current["tag_name"])

    def test_fixture_reproduces_observed_missing_tag_reset(self):
        api = Api(release())
        api.request("PATCH", "/releases/7", {"name": "GridGrind 0.77.0"})
        self.assertEqual("untagged-fixture", api.current["tag_name"])

    def test_wrong_tag_commit_prevents_every_mutation(self):
        api = Api()
        api.sha = "2" * 40
        with self.assertRaisesRegex(publication.PublicationError, "verified commit"):
            self.helper(api).run("prepare")
        self.assertEqual([], api.mutations)
        self.assertEqual([], self.uploads)

    def test_annotated_tag_is_peeled(self):
        api = Api()
        api.annotated = True
        self.helper(api).run("prepare")
        self.assertTrue(api.current["draft"])

    def test_public_prerelease_is_read_only_when_resuming_prepare_or_expose(self):
        api = Api(release(draft=False))
        helper = self.helper(api)
        helper.run("prepare")
        helper.run("expose")
        self.assertEqual([], api.mutations)
        self.assertEqual([], self.uploads)

    def test_public_prerelease_mismatch_never_overwrites_assets_or_notes(self):
        for current, assets in ((release(False, body="changed"), [remote_asset()]),
                                (release(False), [])):
            api = Api(current, assets)
            with self.assertRaises(publication.PublicationError):
                self.helper(api).run("prepare")
            self.assertEqual([], api.mutations)
            self.assertEqual([], self.uploads)

    def test_missing_remote_digest_is_not_accepted(self):
        asset = remote_asset()
        del asset["digest"]
        api = Api(release(False), [asset])
        with self.assertRaisesRegex(publication.PublicationError, "digest"):
            self.helper(api).run("expose")
        self.assertEqual([], api.mutations)

    def test_expose_publishes_only_verified_prerelease_not_latest(self):
        api = Api(release())
        self.helper(api).run("expose")
        self.assertFalse(api.current["draft"])
        self.assertTrue(api.current["prerelease"])
        self.assertEqual("false", api.current["make_latest"])
        self.assertEqual([], self.uploads)

    def test_finalize_requires_previously_exposed_sources(self):
        for current in (None, release()):
            api = Api(current)
            with self.assertRaises(publication.PublicationError):
                self.helper(api).run("finalize")
            self.assertEqual([], api.mutations)

    def test_finalize_promotes_verified_prerelease_to_latest(self):
        api = Api(release(False))
        self.helper(api).run("finalize")
        self.assertFalse(api.current["draft"])
        self.assertFalse(api.current["prerelease"])
        self.assertEqual("true", api.current["make_latest"])
        self.assertEqual([], self.uploads)

    def test_existing_stable_release_always_aborts_read_only(self):
        for stage in ("prepare", "expose", "finalize"):
            api = Api(release(False, False))
            with self.assertRaisesRegex(publication.PublicationError, "manual postflight"):
                self.helper(api).run(stage)
            self.assertEqual([], api.mutations)
            self.assertEqual([], self.uploads)

    def test_api_error_is_not_mistaken_for_absence(self):
        api = Api()
        api.lookup_error = 500
        with self.assertRaises(publication.ApiError):
            self.helper(api).run("prepare")
        self.assertEqual([], api.mutations)

    def test_local_asset_digest_is_independent_of_remote_metadata(self):
        with tempfile.TemporaryDirectory(prefix="gridgrind-release-assets-") as directory:
            path = Path(directory) / "gridgrind.jar"
            path.write_bytes(b"abc")
            assets = publication.local_assets([str(path)])
            self.assertEqual(DIGEST, assets[0].digest)
            self.assertEqual(3, assets[0].size)

    def test_http_retries_get_only_never_ambiguous_mutations(self):
        for method, expected_calls in (("GET", 3), ("POST", 1), ("PATCH", 1)):
            api = publication.Github("owner/GridGrind", "fixture-token")
            api.opener = mock.Mock()
            api.opener.open.side_effect = [urllib.error.HTTPError(
                "https://api.github.com/fixture", 503, "unavailable", {}, io.BytesIO(b"{}")
            ) for _ in range(expected_calls)]
            with mock.patch.object(publication.time, "sleep"), self.assertRaises(publication.ApiError):
                api.request(method, "/releases", None if method == "GET" else {})
            self.assertEqual(expected_calls, api.opener.open.call_count)

    def test_unexpected_draft_assets_require_review(self):
        api = Api(release(), [{"name": "unexpected.zip"}])
        with self.assertRaisesRegex(publication.PublicationError, "unexpected assets"):
            self.helper(api).run("prepare")
        self.assertEqual([], api.mutations)
        self.assertEqual([], self.uploads)


if __name__ == "__main__":
    unittest.main()
