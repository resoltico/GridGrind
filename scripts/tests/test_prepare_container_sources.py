"""Pure source-collection contract tests; no Docker, apt, network or repository writes."""

import hashlib
import importlib.util
import io
import os
from pathlib import Path
import tarfile
import tempfile
import unittest


SPEC = importlib.util.spec_from_file_location(
    "prepare_container_sources", Path(__file__).resolve().parents[1] / "prepare-container-sources.py"
)
collector = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(collector)

PAYLOAD = b"abc"
PAYLOAD_SHA256 = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
PAYLOAD_NAME = "demo_1.2.3.orig.tar.xz"
DSC_NAME = "demo_1.2.3.dsc"
SOURCE_CONFIG = """Types: deb
# http://snapshot.debian.org/archive/debian/20261020T213255Z
URIs: https://deb.debian.org/debian
Suites: trixie trixie-updates
Components: main
Signed-By: /usr/share/keyrings/debian-archive-keyring.gpg
"""


def dsc_text(version="1.2.3", payload_sha256=PAYLOAD_SHA256):
    return (
        "Format: 3.0 (quilt)\nSource: demo\nVersion: " + version
        + "\nChecksums-Sha256:\n " + payload_sha256 + " 3 " + PAYLOAD_NAME + "\n"
    )


class SourceMetadataTest(unittest.TestCase):
    def test_exact_version_and_repeated_matching_archive_stanzas(self):
        text = (
            "Package: demo\nVersion: 1:1.2.3-2\nChecksums-Sha256:\n "
            + "a" * 64 + " 12 " + DSC_NAME + "\n "
            + PAYLOAD_SHA256 + " 3 " + PAYLOAD_NAME + "\n"
        )
        expected = {
            DSC_NAME: {"sha256": "a" * 64, "size": 12},
            PAYLOAD_NAME: {"sha256": PAYLOAD_SHA256, "size": 3},
        }
        self.assertEqual(expected, collector.source_record(text + "\n" + text, "demo", "1:1.2.3-2"))
        with self.assertRaisesRegex(ValueError, "No authenticated source metadata"):
            collector.source_record(text, "demo", "1.2.3-2")
        with self.assertRaisesRegex(ValueError, "Conflicting source archives"):
            collector.source_record(text + "\n" + text.replace(PAYLOAD_SHA256, "b" * 64), "demo", "1:1.2.3-2")

    def test_clearsigned_dsc_and_control_continuation(self):
        signed = (
            "-----BEGIN PGP SIGNED MESSAGE-----\nHash: SHA256\n\n"
            + dsc_text()
            + "Description: source metadata\n -continuation with dash\n"
            + "-----BEGIN PGP SIGNATURE-----\nsignature fixture\n"
        )
        record = collector.controls(signed)[0]
        self.assertEqual("demo", record["Source"])
        self.assertEqual("1.2.3", record["Version"])
        self.assertEqual("source metadata\n-continuation with dash", record["Description"])
        self.assertEqual({PAYLOAD_NAME: {"sha256": PAYLOAD_SHA256, "size": 3}},
                         collector.checksums(record["Checksums-Sha256"]))

    def test_unsafe_or_malformed_checksum_rows_fail(self):
        rows = [
            PAYLOAD_SHA256 + " 3 ../outside.tar.xz",
            PAYLOAD_SHA256 + " 3 /absolute.tar.xz",
            PAYLOAD_SHA256[:-1] + " 3 demo.tar.xz",
            PAYLOAD_SHA256 + " -3 demo.tar.xz",
            PAYLOAD_SHA256 + " 3 demo.tar.xz\n" + PAYLOAD_SHA256 + " 3 demo.tar.xz",
        ]
        for row in rows:
            with self.subTest(row=row), self.assertRaises(ValueError):
                collector.checksums(row)


class AptConfigurationTest(unittest.TestCase):
    def test_snapshot_conversion_preserves_signature_configuration(self):
        converted = collector.source_config(SOURCE_CONFIG, "debian.sources")
        self.assertEqual(2, converted.count("Types: deb-src"))
        self.assertEqual(2, converted.count("Signed-By: /usr/share/keyrings/debian-archive-keyring.gpg"))
        self.assertEqual(1, converted.count("Check-Valid-Until: no"))
        self.assertIn("URIs: https://deb.debian.org/debian", converted)
        self.assertIn("URIs: http://snapshot.debian.org/archive/debian/20261020T213255Z", converted)

    def test_authentication_bypasses_are_rejected(self):
        for setting in ["Trusted", "Allow-Insecure", "Allow-Unauthenticated", "Allow-Weak", "Allow-Downgrade-To-Insecure"]:
            for truth in ["yes", "true", "1"]:
                with self.subTest(setting=setting, truth=truth), self.assertRaisesRegex(ValueError, "Unauthenticated"):
                    collector.source_config(SOURCE_CONFIG + setting + ": " + truth + "\n", "debian.sources")
        with self.assertRaisesRegex(ValueError, "Unauthenticated"):
            collector.source_config("deb [trusted=true] https://deb.debian.org/debian trixie main", "debian.list")

    def test_unsupported_source_uri_and_missing_uri_fail(self):
        for uri in ["file:/tmp/unsigned", "ftp://deb.debian.org/debian", "https://example.com/debian", "https://u:p@deb.debian.org/debian"]:
            for name, text in [
                ("debian.sources", SOURCE_CONFIG.replace("https://deb.debian.org/debian", uri)),
                ("debian.list", "deb " + uri + " trixie main"),
            ]:
                with self.subTest(uri=uri, name=name), self.assertRaises(ValueError):
                    collector.source_config(text, name)
        with self.assertRaisesRegex(ValueError, "no archive URI"):
            collector.source_config("Types: deb\nSuites: trixie\n", "broken.sources")

    def test_apt_uri_output_is_parsed_and_malformed_rows_fail(self):
        text = "Need to get source archives.\n'https://deb.debian.org/pool/demo.dsc' demo.dsc 12 SHA256:abc\n"
        self.assertEqual({"demo.dsc": {"url": "https://deb.debian.org/pool/demo.dsc", "size": 12}},
                         collector.uri_records(text))
        for row in [
            "'https://deb.debian.org/pool/demo.dsc' ../outside 12 SHA256:abc",
            "'https://deb.debian.org/pool/demo.dsc' demo.dsc nope SHA256:abc",
            "'file:/tmp/demo.dsc' demo.dsc 12 SHA256:abc",
            "'https://example.com/demo.dsc' demo.dsc 12 SHA256:abc",
            "No source URIs available",
        ]:
            with self.subTest(row=row), self.assertRaises(ValueError):
                collector.uri_records(row)


class SourcePayloadTest(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        (self.directory / PAYLOAD_NAME).write_bytes(PAYLOAD)
        self.set_dsc(dsc_text())

    def set_dsc(self, text):
        data = text.encode()
        (self.directory / DSC_NAME).write_bytes(data)
        self.expected = {
            DSC_NAME: {"sha256": hashlib.sha256(data).hexdigest(), "size": len(data)},
            PAYLOAD_NAME: {"sha256": PAYLOAD_SHA256, "size": 3},
        }

    def test_valid_package_includes_every_dsc_payload(self):
        collector.verify_package(self.directory, "demo", "1.2.3", self.expected)

    def test_truncated_missing_and_same_length_changed_payload_fail(self):
        for data in [b"ab", b"abd", None]:
            with self.subTest(data=data):
                path = self.directory / PAYLOAD_NAME
                if data is None:
                    path.unlink()
                else:
                    path.write_bytes(data)
                with self.assertRaisesRegex(ValueError, "mismatched source payload"):
                    collector.verify_package(self.directory, "demo", "1.2.3", self.expected)

    def test_wrong_dsc_identity_fails_after_file_hashes_pass(self):
        for field, wrong in [("Source: demo", "Source: other"), ("Version: 1.2.3", "Version: 1.2.4")]:
            self.set_dsc(dsc_text().replace(field, wrong))
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, "Wrong source tuple"):
                collector.verify_package(self.directory, "demo", "1.2.3", self.expected)

    def test_index_and_dsc_payload_hashes_must_agree(self):
        self.set_dsc(dsc_text(payload_sha256="b" * 64))
        with self.assertRaisesRegex(ValueError, "archive index disagree"):
            collector.verify_package(self.directory, "demo", "1.2.3", self.expected)


class ManifestDescriptorTest(unittest.TestCase):
    def test_platform_manifest_digest_is_required(self):
        digest = "sha256:" + "a" * 64
        for media_type in collector.MANIFEST_MEDIA_TYPES:
            self.assertEqual(digest, collector.manifest_digest({"Descriptor": {"mediaType": media_type, "digest": digest}}))
        for info in [
            {"Id": digest},
            {"Descriptor": None},
            {"Descriptor": {"mediaType": "application/vnd.oci.image.index.v1+json", "digest": digest}},
            {"Descriptor": {"mediaType": "application/vnd.docker.distribution.manifest.list.v2+json", "digest": digest}},
            {"Descriptor": {"mediaType": "application/vnd.oci.image.manifest.v1+json", "digest": "sha256:short"}},
            {"Descriptor": {"mediaType": "application/vnd.oci.image.manifest.v1+json", "digest": None}},
        ]:
            with self.subTest(info=info), self.assertRaises(ValueError):
                collector.manifest_digest(info)


class SourceCacheTest(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.cache = self.directory / "prior" / "content" / "debian"
        self.cached_tuple = self.cache / "demo" / "1.2.3"
        self.cached_tuple.mkdir(parents=True)
        data = dsc_text().encode()
        (self.cached_tuple / DSC_NAME).write_bytes(data)
        (self.cached_tuple / PAYLOAD_NAME).write_bytes(PAYLOAD)
        self.expected = {
            DSC_NAME: {"sha256": hashlib.sha256(data).hexdigest(), "size": len(data)},
            PAYLOAD_NAME: {"sha256": PAYLOAD_SHA256, "size": 3},
        }
        self.destination = self.directory / "new-output"
        self.destination.mkdir()

    def test_reuse_has_identical_bytes_and_independent_files(self):
        self.assertTrue(collector.reuse_source_package(self.cache, self.destination, "demo", "1.2.3", self.expected))
        for name in self.expected:
            self.assertEqual((self.cached_tuple / name).read_bytes(), (self.destination / name).read_bytes())
            self.assertNotEqual((self.cached_tuple / name).stat().st_ino, (self.destination / name).stat().st_ino)
        fresh_archive = self.directory / "fresh.tar.gz"
        cached_archive = self.directory / "cached.tar.gz"
        collector.write_archive(self.cached_tuple, fresh_archive, "0.77.0")
        collector.write_archive(self.destination, cached_archive, "0.77.0")
        self.assertEqual(fresh_archive.read_bytes(), cached_archive.read_bytes())
        (self.destination / PAYLOAD_NAME).write_bytes(b"new")
        self.assertEqual(PAYLOAD, (self.cached_tuple / PAYLOAD_NAME).read_bytes())

    def test_missing_tuple_is_not_reused(self):
        self.assertFalse(collector.reuse_source_package(self.cache, self.destination, "demo", "1.2.4", self.expected))
        self.assertEqual([], list(self.destination.iterdir()))

    def test_corrupted_existing_tuple_fails_without_mutation_or_copy(self):
        (self.cached_tuple / PAYLOAD_NAME).write_bytes(b"bad")
        with self.assertRaisesRegex(ValueError, "mismatched source payload"):
            collector.reuse_source_package(self.cache, self.destination, "demo", "1.2.3", self.expected)
        self.assertEqual(b"bad", (self.cached_tuple / PAYLOAD_NAME).read_bytes())
        self.assertEqual([], list(self.destination.iterdir()))

    def test_wrong_version_in_existing_tuple_is_rejected(self):
        wrong_directory = self.cache / "demo" / "1.2.4"
        self.cached_tuple.rename(wrong_directory)
        with self.assertRaisesRegex(ValueError, "Wrong source tuple"):
            collector.reuse_source_package(self.cache, self.destination, "demo", "1.2.4", self.expected)
        self.assertEqual([], list(self.destination.iterdir()))


class SourceArchiveTest(unittest.TestCase):
    def test_build_sources_are_copied_unchanged_with_licensing_and_hashes(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            repository = root / "repository"
            content = root / "content"
            for name in collector.BUILD_INPUTS:
                source = repository / name
                source.parent.mkdir(parents=True, exist_ok=True)
                source.write_bytes(("source fixture: " + name + "\n").encode())
            records = collector.copy_build_inputs(repository, content)
            self.assertEqual(set(collector.BUILD_INPUTS), {record["sourcePath"] for record in records})
            for record in records:
                copied = content / record["archivePath"]
                original = repository / record["sourcePath"]
                self.assertEqual(original.read_bytes(), copied.read_bytes())
                self.assertEqual(hashlib.sha256(original.read_bytes()).hexdigest(), record["sha256"])
                self.assertEqual(original.stat().st_size, record["size"])

    def test_output_path_and_input_metadata_do_not_change_archive_bytes(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            first = root / "first-input"
            second = root / "other-input"
            for directory, timestamp, mode in [(first, 123456, 0o600), (second, 999999, 0o755)]:
                (directory / "vendor").mkdir(parents=True)
                # These stand in for already-compressed upstream bytes: never rewrite their contents.
                payload = directory / "vendor" / "source.tar.gz"
                payload.write_bytes(b"upstream archive bytes\x00\xff")
                payload.chmod(mode)
                os.utime(payload, (timestamp, timestamp))
                os.utime(directory / "vendor", (timestamp, timestamp))
                os.utime(directory, (timestamp, timestamp))
            one = root / "first-output.tar.gz"
            two = root / "another-output-name.tar.gz"
            collector.write_archive(first, one, "0.77.0")
            collector.write_archive(second, two, "0.77.0")
            self.assertEqual(one.read_bytes(), two.read_bytes())
            raw = one.read_bytes()
            self.assertEqual(0, int.from_bytes(raw[4:8], "little"), "Gzip timestamp must be fixed")
            self.assertEqual(0, raw[3] & 8, "Gzip must not carry the output filename")
            with tarfile.open(fileobj=io.BytesIO(raw), mode="r:gz") as archive:
                for member in archive.getmembers():
                    self.assertEqual((0, 0, "root", "root", 0),
                                     (member.uid, member.gid, member.uname, member.gname, member.mtime))
                    self.assertEqual(0o755 if member.isdir() else 0o644, member.mode)
                payload = archive.extractfile("gridgrind-container-sources-0.77.0/vendor/source.tar.gz")
                self.assertEqual(b"upstream archive bytes\x00\xff", payload.read())


if __name__ == "__main__":
    unittest.main()
