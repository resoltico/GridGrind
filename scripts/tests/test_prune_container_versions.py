"""Pure fixture checks: no network or package deletions are performed."""

import contextlib
import hashlib
import io
import importlib.util
import json
from pathlib import Path
import sys
import unittest
from unittest import mock

SPEC = importlib.util.spec_from_file_location(
    "container_retention", Path(__file__).parents[1] / "prune-container-versions.py"
)
prune = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = prune
SPEC.loader.exec_module(prune)


class Registry:
    def __init__(self):
        self.bodies = {}

    def add(self, label, children=None):
        document = {"schemaVersion": 2, "annotations": {"fixture": label}}
        if children is None:
            document.update({
                "mediaType": "application/vnd.oci.image.manifest.v1+json",
                "config": {"digest": "sha256:" + "0" * 64}, "layers": [],
            })
        else:
            document.update({
                "mediaType": "application/vnd.oci.image.index.v1+json",
                "manifests": [{"digest": child} for child in children],
            })
        raw = json.dumps(document, sort_keys=True).encode()
        name = "sha256:" + hashlib.sha256(raw).hexdigest()
        self.bodies[name] = raw
        return name

    def manifest(self, name):
        if name not in self.bodies:
            raise prune.MissingManifest("Fixture manifest is unavailable")
        return self.bodies[name]


class Api:
    def __init__(self, versions):
        self.current = versions
        self.deleted = []

    def versions(self):
        return self.current

    def delete(self, version):
        self.deleted.append(version.identifier)


def version(identifier, name, tags=(), date="2026-09-01T00:00:00Z"):
    return prune.Version(identifier, name, tuple(tags), date)


class RetentionTest(unittest.TestCase):
    def setUp(self):
        self.registry = Registry()
        self.versions = []
        self.children = []
        self.roots = []
        for number in range(1, 7):
            child = self.registry.add("platform-%s" % number)
            root = self.registry.add("release-%s" % number, [child])
            self.children.append(child)
            self.roots.append(root)
            self.versions.extend([
                version(number, root, ["0.%s.0" % number], "2026-10-%02dT00:00:00Z" % number),
                # Children precede their tagged index, including retained releases.
                version(100 + number, child),
            ])

    def plan(self):
        return prune.make_plan(self.versions, prune.ManifestGraph(self.registry))

    def test_child_older_than_retained_root_survives(self):
        plan = self.plan()
        self.assertEqual({1, 101}, {item.identifier for item in plan.deletions})
        self.assertEqual({2, 3, 4, 5, 6}, {item.identifier for item in plan.retained})
        api = Api(self.versions)
        prune.apply_plan(api, plan)
        self.assertEqual([1, 101], api.deleted)

    def test_shared_child_is_protected_by_retained_root(self):
        root = self.registry.add("retired-with-shared-child", [self.children[0], self.children[5]])
        self.versions[0] = version(1, root, ["0.1.0"], "2026-10-01T00:00:00Z")
        deletions = {item.identifier for item in self.plan().deletions}
        self.assertEqual({1, 101}, deletions)
        self.assertNotIn(106, deletions)

    def test_unknown_untagged_orphan_is_retained(self):
        orphan = self.registry.add("orphan")
        self.versions.append(version(999, orphan))
        self.assertNotIn(999, {item.identifier for item in self.plan().deletions})

    def test_nonstable_alias_protects_an_otherwise_retired_root(self):
        self.versions[0] = version(1, self.roots[0], ["0.1.0", "latest"], "2026-10-01T00:00:00Z")
        self.assertEqual((), self.plan().deletions)

    def test_runtime_and_source_pair_consume_one_stable_anchor(self):
        for number in range(1, 7):
            source = self.registry.add("sources-%s" % number)
            self.versions.append(version(200 + number, source, ["0.%s.0-sources" % number]))
        plan = self.plan()
        self.assertEqual({2, 3, 4, 5, 6}, {item.identifier for item in plan.retained})
        self.assertEqual({1, 101, 201}, {item.identifier for item in plan.deletions})
        self.assertEqual({201}, {item.identifier for item in plan.retired_sources})

    def test_retired_source_graph_follows_its_retired_runtime_group(self):
        child = self.registry.add("old-source-platform")
        root = self.registry.add("old-source-index", [child])
        self.versions.extend([version(201, root, ["0.1.0-sources"]), version(301, child)])
        self.assertEqual({1, 101, 201, 301}, {item.identifier for item in self.plan().deletions})

    def test_incoming_source_without_runtime_anchor_is_protected(self):
        child = self.registry.add("incoming-source-platform")
        root = self.registry.add("incoming-source-index", [child])
        self.versions.extend([version(207, root, ["0.7.0-sources"]), version(307, child)])
        self.assertEqual({1, 101}, {item.identifier for item in self.plan().deletions})

    def test_shared_source_child_is_protected_by_retained_source_root(self):
        shared = self.registry.add("shared-source-platform")
        retired = self.registry.add("old-source-index", [shared])
        current = self.registry.add("current-source-index", [shared])
        self.versions.extend([
            version(201, retired, ["0.1.0-sources"]),
            version(206, current, ["0.6.0-sources"]), version(300, shared),
        ])
        self.assertEqual({1, 101, 201}, {item.identifier for item in self.plan().deletions})

    def test_retired_source_with_other_alias_is_still_protected(self):
        source = self.registry.add("aliased-source")
        self.versions.append(version(201, source, ["0.1.0-sources", "source-preview"]))
        self.assertEqual({1, 101}, {item.identifier for item in self.plan().deletions})

    def test_other_tagged_root_protects_its_references(self):
        preview = self.registry.add("preview", [self.children[0]])
        self.versions.append(version(900, preview, ["preview"]))
        self.assertEqual({1}, {item.identifier for item in self.plan().deletions})

    def test_retired_child_404_only_deletes_resolved_index_and_preserves_shared_child(self):
        retired = self.registry.add("retired-shares-current", [self.children[0], self.children[5]])
        self.versions[0] = version(1, retired, ["0.1.0"], "2026-10-01T00:00:00Z")
        del self.registry.bodies[self.children[0]]
        plan = self.plan()
        self.assertEqual({1}, {item.identifier for item in plan.deletions})
        self.assertEqual((self.children[0],), plan.unavailable_retired)
        api = Api(self.versions)
        prune.apply_plan(api, plan)
        self.assertEqual([1], api.deleted)

    def test_retired_root_404_proves_no_candidates(self):
        del self.registry.bodies[self.roots[0]]
        plan = self.plan()
        self.assertEqual((), plan.deletions)
        self.assertEqual((self.roots[0],), plan.unavailable_retired)

    def test_retired_403_or_server_error_still_causes_no_deletes(self):
        for status in (403, 500):
            api = Api(self.versions)
            original = self.registry.manifest
            def response(name):
                if name == self.children[0]:
                    raise prune.HttpFailure("GET", status)
                return original(name)
            with mock.patch.object(self.registry, "manifest", side_effect=response):
                with self.assertRaises(prune.HttpFailure):
                    prune.apply_plan(api, self.plan())
            self.assertEqual([], api.deleted)

    def test_registry_distinguishes_explicit_404_from_other_errors(self):
        for status in (404, 403, 500):
            http = mock.Mock()
            http.request.side_effect = [b'{"token":"fixture"}', prune.HttpFailure("GET", status)]
            registry = prune.Registry(http, "owner", "package")
            expected = prune.MissingManifest if status == 404 else prune.HttpFailure
            with self.assertRaises(expected):
                registry.manifest(self.roots[0])

    def test_missing_protected_manifest_also_causes_no_deletes(self):
        del self.registry.bodies[self.children[5]]
        api = Api(self.versions)
        with self.assertRaisesRegex(prune.PruningError, "unavailable"):
            prune.apply_plan(api, self.plan())
        self.assertEqual([], api.deleted)

    def test_manifest_hash_mismatch_causes_no_deletes(self):
        self.registry.bodies[self.roots[0]] += b" "
        api = Api(self.versions)
        with self.assertRaisesRegex(prune.PruningError, "do not match"):
            prune.apply_plan(api, self.plan())
        self.assertEqual([], api.deleted)

    def test_state_change_aborts_before_first_delete(self):
        plan = self.plan()
        api = Api(self.versions + [version(999, self.registry.add("new"), ["new-tag"])])
        with self.assertRaisesRegex(prune.PruningError, "state changed"):
            prune.apply_plan(api, plan)
        self.assertEqual([], api.deleted)

    def test_fewer_than_retention_count_needs_no_manifest_requests(self):
        versions = [version(1, "sha256:" + "1" * 64, ["0.1.0"])]
        self.assertEqual((), prune.make_plan(versions, prune.ManifestGraph(self.registry)).deletions)

    def test_duplicate_paginated_ids_are_not_a_valid_snapshot(self):
        with self.assertRaisesRegex(prune.PruningError, "duplicate IDs"):
            prune.snapshot(self.versions + [self.versions[0]])

    def test_default_command_is_dry_run(self):
        api = Api(self.versions)
        with mock.patch.object(prune, "Github", return_value=api), mock.patch.object(
            prune, "Registry", return_value=self.registry
        ), mock.patch.dict(prune.os.environ, {"GH_TOKEN": "fixture-token"}), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(0, prune.main(["--owner", "fixture-owner"]))
        self.assertEqual([], api.deleted)

    def test_apply_command_requires_explicit_flag(self):
        api = Api(self.versions)
        with mock.patch.object(prune, "Github", return_value=api), mock.patch.object(
            prune, "Registry", return_value=self.registry
        ), mock.patch.dict(prune.os.environ, {"GH_TOKEN": "fixture-token"}), contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(0, prune.main(["--owner", "fixture-owner", "--apply"]))
        self.assertEqual([1, 101], api.deleted)

    def test_redirect_cannot_forward_credentials(self):
        with self.assertRaisesRegex(prune.PruningError, "redirect"):
            prune.NoRedirect().redirect_request(None, None, 302, None, None, "https://elsewhere")

    def test_unbounded_or_nonsha_manifest_digest_is_rejected(self):
        for value in ("sha256:123", "https://elsewhere/manifest", "sha512:" + "0" * 128):
            with self.assertRaises(prune.PruningError):
                prune.digest(value)


class PendingReleaseTest(unittest.TestCase):
    def test_new_runtime_reserves_one_future_slot(self):
        self.assertEqual(4, prune.preflight_keep([], 5, "0.77.0"))

    def test_existing_runtime_keeps_all_current_slots(self):
        rows = [version(1, "sha256:" + "a" * 64, ["0.77.0"])]
        self.assertEqual(5, prune.preflight_keep(rows, 5, "0.77.0"))

    def test_source_only_does_not_consume_a_runtime_slot(self):
        rows = [version(1, "sha256:" + "a" * 64, ["0.77.0-sources"])]
        self.assertEqual(4, prune.preflight_keep(rows, 5, "0.77.0"))


if __name__ == "__main__":
    unittest.main()
