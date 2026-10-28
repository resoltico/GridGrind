import hashlib
import importlib.util
import io
import json
from pathlib import Path
import unittest
from unittest import mock
import urllib.request

spec = importlib.util.spec_from_file_location("release_image_verifier", Path(__file__).parents[1] / "verify-container-release.py")
v = importlib.util.module_from_spec(spec)
spec.loader.exec_module(v)


class VerificationTest(unittest.TestCase):
    def setUp(self):
        self.index = {"manifests": []}
        self.nodes = {}
        self.blobs = {}
        self.runtime = {"amd64": "sha256:" + "a" * 64, "arm64": "sha256:" + "b" * 64}
        self.source = {"images": [{"platform": "linux/amd64", "manifestDigest": self.runtime["amd64"]}, {"platform": "linux/arm64", "manifestDigest": self.runtime["arm64"]}]}
        self.attestation = {"amd64": "sha256:" + "c" * 64, "arm64": "sha256:" + "d" * 64}
        for arch, config in [("amd64", "sha256:" + "e" * 64), ("arm64", "sha256:" + "f" * 64)]:
            digest = self.runtime[arch]
            self.index["manifests"].append({"platform": {"os": "linux", "architecture": arch}, "digest": digest})
            self.index["manifests"].append({"platform": {"os": "unknown", "architecture": "unknown"}, "digest": self.attestation[arch], "annotations": {"vnd.docker.reference.type": "attestation-manifest", "vnd.docker.reference.digest": digest}})
            self.nodes[digest] = {"config": {"digest": config}}
            self.nodes[self.attestation[arch]] = {"layers": []}
            for kind, predicate in [
                ("https://slsa.dev/provenance/v1", {"buildDefinition": {"buildType": "https://example.test/build"}}),
                (v.SPDX_TYPE, {"spdxVersion": "SPDX-2.3", "packages": [{"SPDXID": "SPDXRef-package", "name": "package"}]}),
            ]:
                statement = {"_type": "https://in-toto.io/Statement/v1", "predicateType": kind,
                             "subject": [{"name": "image", "digest": {"sha256": digest[7:]}}], "predicate": predicate}
                self.nodes[self.attestation[arch]]["layers"].append(self.layer(statement))

    def layer(self, statement):
        return self.raw_layer(json.dumps(statement).encode(), statement["predicateType"])

    def raw_layer(self, raw, kind):
        digest = "sha256:" + hashlib.sha256(raw).hexdigest()
        self.blobs[digest] = raw
        return {"digest": digest, "size": len(raw), "mediaType": "application/vnd.in-toto+json",
                "annotations": {"in-toto.io/predicate-type": kind}}

    def fetch_blob(self, layer):
        return v.verify_blob(self.blobs[layer["digest"]], layer)

    def check(self):
        return v.verify(self.index, self.nodes.__getitem__, self.source, self.fetch_blob)

    def change_statement(self, mutation, layer_index=0):
        layers = self.nodes[self.attestation["amd64"]]["layers"]
        statement = json.loads(self.blobs[layers[layer_index]["digest"]])
        mutation(statement)
        layers[layer_index] = self.layer(statement)

    def test_both_manifest_bindings_and_actual_statements(self):
        result = self.check()
        self.assertEqual(2, len(result))
        self.assertEqual(2, len(result[0]["statements"]))
        self.assertEqual("sha256:" + "e" * 64, result[0]["configDigest"])
        self.assertEqual(1, result[0]["statements"][1]["packageCount"])

    def test_in_toto_v01_and_slsa_v02_fields(self):
        def replace(statement):
            statement.update({"_type": "https://in-toto.io/Statement/v0.1", "predicateType": "https://slsa.dev/provenance/v0.2", "predicate": {"buildType": "buildkit"}})
        self.change_statement(replace)
        self.assertEqual(2, len(self.check()))

    def test_different_published_manifest_is_rejected(self):
        self.source["images"][1]["manifestDigest"] = "sha256:" + "0" * 64
        with self.assertRaisesRegex(ValueError, "differs"):
            self.check()

    def test_changed_config_cannot_pass_unchanged_manifest_digest(self):
        original = json.dumps({"config": {"digest": "sha256:" + "e" * 64}}).encode()
        changed = json.dumps({"config": {"digest": "sha256:" + "f" * 64}}).encode()
        registry = object.__new__(v.Registry)
        registry.repository, registry.token, registry.metadata = "owner/image", "fixture", mock.Mock()
        with mock.patch.object(v, "read", return_value=changed), self.assertRaisesRegex(ValueError, "Manifest size/digest mismatch"):
            registry.fetch("sha256:" + hashlib.sha256(original).hexdigest())

    def test_missing_sbom_is_rejected(self):
        self.nodes[self.attestation["amd64"]]["layers"].pop()
        with self.assertRaisesRegex(ValueError, "SBOM"):
            self.check()

    def test_missing_platform_is_rejected(self):
        self.index["manifests"] = self.index["manifests"][:2]
        with self.assertRaisesRegex(ValueError, "platform"):
            self.check()

    def test_wrong_subject_is_rejected_after_valid_blob_hash(self):
        self.change_statement(lambda statement: statement.update({"subject": [{"digest": {"sha256": "f" * 64}}]}))
        with self.assertRaisesRegex(ValueError, "subject"):
            self.check()

    def test_annotation_mismatch_is_rejected(self):
        self.nodes[self.attestation["amd64"]]["layers"][0]["annotations"]["in-toto.io/predicate-type"] = v.SPDX_TYPE
        with self.assertRaisesRegex(ValueError, "annotation mismatch"):
            self.check()

    def test_invalid_json_payload_is_rejected_after_valid_blob_hash(self):
        self.nodes[self.attestation["amd64"]]["layers"][0] = self.raw_layer(b"not JSON", "https://slsa.dev/provenance/v1")
        with self.assertRaisesRegex(ValueError, "Invalid attestation JSON"):
            self.check()

    def test_empty_slsa_build_type_is_rejected(self):
        self.change_statement(lambda statement: statement.update({"predicate": {"buildDefinition": {"buildType": " "}}}))
        with self.assertRaisesRegex(ValueError, "build type"):
            self.check()

    def test_empty_spdx_package_list_is_rejected(self):
        self.change_statement(lambda statement: statement["predicate"].update({"packages": []}), 1)
        with self.assertRaisesRegex(ValueError, "package records"):
            self.check()

    def test_missing_spdx_version_is_rejected(self):
        self.change_statement(lambda statement: statement["predicate"].pop("spdxVersion"), 1)
        with self.assertRaisesRegex(ValueError, "SPDX version"):
            self.check()

    def test_blob_hash_mismatch_is_rejected(self):
        layer = self.nodes[self.attestation["amd64"]]["layers"][0]
        raw = self.blobs[layer["digest"]]
        self.blobs[layer["digest"]] = b"[" + raw[1:]
        with self.assertRaisesRegex(ValueError, "hash mismatch"):
            self.check()

    def test_oversized_descriptor_rejected_before_fetch(self):
        registry = object.__new__(v.Registry)
        registry.blobs = mock.Mock()
        layer = dict(self.nodes[self.attestation["amd64"]]["layers"][0], size=v.BLOB_LIMIT + 1)
        with self.assertRaisesRegex(ValueError, "oversized"):
            registry.fetch_blob(layer)
        registry.blobs.open.assert_not_called()

    def test_oversized_response_body_is_bounded(self):
        response = io.BytesIO(b"12345")
        response.headers = {}
        opener = mock.Mock()
        opener.open.return_value = response
        with self.assertRaisesRegex(ValueError, "Oversized registry response"):
            v.read(opener, "https://example.test", 4)

    def test_blob_redirect_strips_credentials_even_on_same_host(self):
        for destination in ("https://cdn.example.test/blob?signature=fixture", "https://ghcr.io/blob"):
            request = urllib.request.Request("https://ghcr.io/v2/owner/image/blobs/sha256:fixture", headers={"Authorization": "Bearer fixture", "Proxy-Authorization": "fixture", "Cookie": "fixture", "Accept": "application/json"})
            request.add_unredirected_header("Authorization", "Bearer unredirected-fixture")
            redirected = v.BlobRedirect().redirect_request(request, None, 307, "redirect", {}, destination)
            self.assertFalse(any(name.lower() in {"authorization", "proxy-authorization", "cookie"} for name, _ in redirected.header_items()))
            self.assertEqual("application/json", redirected.get_header("Accept"))

    def test_blob_redirect_rejects_http_and_embedded_credentials(self):
        request = urllib.request.Request("https://ghcr.io/blob")
        for destination in ("http://cdn.example.test/blob", "https://user:secret@cdn.example.test/blob"):
            with self.assertRaisesRegex(ValueError, "HTTPS without embedded"):
                v.BlobRedirect().redirect_request(request, None, 307, "redirect", {}, destination)

    def test_metadata_redirect_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "metadata must not redirect"):
            v.NoRedirect().redirect_request(None)


if __name__ == "__main__":
    unittest.main()
