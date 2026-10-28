#!/usr/bin/env python3
"""Check published runtime manifests and attestation content against inspected source inputs.

These checks establish content/subject bindings and required predicate fields;
they do not establish a trusted builder identity or a SLSA assurance level.
"""
import argparse
import hashlib
import json
from pathlib import Path
import re
import time
import urllib.error
import urllib.parse
import urllib.request

DIGEST = re.compile(r"sha256:[0-9a-f]{64}\Z")
ACCEPT = ", ".join(("application/vnd.oci.image.index.v1+json", "application/vnd.oci.image.manifest.v1+json",
                    "application/vnd.docker.distribution.manifest.list.v2+json", "application/vnd.docker.distribution.manifest.v2+json"))
LIMIT = 4 * 1024 * 1024
BLOB_LIMIT = 64 * 1024 * 1024
SLSA_TYPES = {"https://slsa.dev/provenance/v0.2", "https://slsa.dev/provenance/v1"}
SPDX_TYPE = "https://spdx.dev/Document"
STATEMENT_TYPES = {"https://in-toto.io/Statement/v0.1", "https://in-toto.io/Statement/v1"}


def require(value, message):
    if not value:
        raise ValueError(message)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *_):
        raise ValueError("Registry metadata must not redirect bearer credentials")


class BlobRedirect(urllib.request.HTTPRedirectHandler):
    max_redirections = 5
    max_repeats = 2

    def redirect_request(self, request, response, code, message, headers, new_url):
        target = urllib.parse.urljoin(request.full_url, new_url)
        parsed = urllib.parse.urlsplit(target)
        require(parsed.scheme == "https" and parsed.hostname and parsed.username is None and parsed.password is None,
                "Blob redirect must use HTTPS without embedded credentials")
        redirected = super().redirect_request(request, response, code, message, headers, target)
        if redirected is not None:
            for collection in (redirected.headers, redirected.unredirected_hdrs):
                for name in list(collection):
                    if name.lower() in {"authorization", "proxy-authorization", "cookie"}:
                        del collection[name]
        return redirected


def read(opener, request, limit):
    for attempt in range(3):
        try:
            with opener.open(request, timeout=30) as response:
                length = response.headers.get("Content-Length")
                require(length is None or length.isdecimal() and int(length) <= limit,
                        "Invalid or oversized registry response Content-Length")
                raw = response.read(limit + 1)
            require(len(raw) <= limit, "Oversized registry response")
            return raw
        except urllib.error.HTTPError as error:
            status = error.code
            error.close()
            if status not in {429, 500, 502, 503, 504} or attempt == 2:
                raise ValueError("Registry read failed with HTTP " + str(status)) from None
        except (urllib.error.URLError, TimeoutError, OSError):
            if attempt == 2:
                # Never expose signed redirect URLs or bearer credentials in diagnostics.
                raise ValueError("Registry read failed or timed out") from None
        time.sleep(2 ** attempt)
    raise ValueError("Registry read did not complete")


def blob_descriptor(descriptor):
    require(isinstance(descriptor, dict), "Invalid attestation blob descriptor")
    name, size = descriptor.get("digest"), descriptor.get("size")
    require(isinstance(name, str) and DIGEST.fullmatch(name), "Invalid attestation blob digest")
    require(type(size) is int and 0 < size <= BLOB_LIMIT, "Invalid or oversized attestation blob size")
    require(descriptor.get("mediaType") == "application/vnd.in-toto+json", "Unsupported attestation blob media type")
    return name, size


def verify_blob(raw, descriptor):
    name, size = blob_descriptor(descriptor)
    require(len(raw) <= BLOB_LIMIT and len(raw) == size, "Attestation blob size mismatch")
    require("sha256:" + hashlib.sha256(raw).hexdigest() == name, "Attestation blob hash mismatch")
    try:
        return json.loads(raw)
    except (ValueError, UnicodeError):
        raise ValueError("Invalid attestation JSON") from None


def verify_statement(statement, descriptor, runtime_digest):
    require(isinstance(statement, dict), "Attestation payload is not a statement")
    require(statement.get("_type") in STATEMENT_TYPES, "Unsupported in-toto statement type")
    predicate_type = statement.get("predicateType")
    annotation = descriptor.get("annotations", {}).get("in-toto.io/predicate-type")
    require(isinstance(predicate_type, str) and predicate_type == annotation, "Attestation predicate annotation mismatch")
    subjects = statement.get("subject")
    require(isinstance(subjects, list) and subjects, "Attestation statement lacks subjects")
    require(any(isinstance(subject, dict) and isinstance(subject.get("digest"), dict)
                and subject["digest"].get("sha256") == runtime_digest.removeprefix("sha256:")
                for subject in subjects), "Attestation subject does not bind the runtime platform manifest")
    predicate = statement.get("predicate")
    require(isinstance(predicate, dict) and predicate, "Attestation predicate is empty or invalid")
    details = {}
    if predicate_type in SLSA_TYPES:
        definition = predicate.get("buildDefinition", {}) if predicate_type.endswith("/v1") else predicate
        require(isinstance(definition, dict) and isinstance(definition.get("buildType"), str)
                and definition["buildType"].strip(), "SLSA predicate lacks a nonempty build type")
        details["buildTypePresent"] = True
    elif predicate_type == SPDX_TYPE:
        version = predicate.get("spdxVersion")
        packages = predicate.get("packages")
        require(isinstance(version, str) and re.fullmatch(r"SPDX-[0-9]+\.[0-9]+", version), "SPDX predicate lacks an SPDX version")
        require(isinstance(packages, list) and packages and all(
            isinstance(package, dict) and isinstance(package.get("name"), str) and package["name"].strip()
            for package in packages), "SPDX predicate lacks nonempty package records")
        details.update({"spdxVersion": version, "packageCount": len(packages)})
    else:
        raise ValueError("Unsupported attestation predicate type")
    return {"predicateType": predicate_type, "statementType": statement["_type"],
            "blobDigest": descriptor["digest"], "blobSize": descriptor["size"], **details}


def verify(index, fetch, source, fetch_blob):
    expected = {row["platform"]: row["manifestDigest"] for row in source["images"]}
    require(len(source["images"]) == 2 and set(expected) == {"linux/amd64", "linux/arm64"}, "Source mapping must name both platforms exactly once")
    require(all(isinstance(value, str) and DIGEST.fullmatch(value) for value in expected.values()), "Source mapping requires typed runtime manifest digests")
    ordinary = {}
    attestations = {}
    for entry in index.get("manifests", []):
        platform = entry.get("platform", {})
        key = platform.get("os", "") + "/" + platform.get("architecture", "")
        annotations = entry.get("annotations", {})
        if annotations.get("vnd.docker.reference.type") == "attestation-manifest":
            target = annotations.get("vnd.docker.reference.digest")
            require(target not in attestations, "Repeated attestation target")
            attestations[target] = entry["digest"]
        else:
            require(key in expected and key not in ordinary, "Unexpected or repeated runtime platform")
            ordinary[key] = entry["digest"]
    require(set(ordinary) == set(expected), "Published image lacks a required runtime platform")
    require(set(attestations) == set(ordinary.values()), "Unexpected or missing attestation target")
    results = []
    for platform, digest in sorted(ordinary.items()):
        require(isinstance(digest, str) and DIGEST.fullmatch(digest), "Invalid runtime manifest digest")
        require(digest == expected[platform], "Published runtime manifest differs from the tested/source-inventoried image: " + platform)
        manifest = fetch(digest)
        config_digest = manifest.get("config", {}).get("digest")
        require(isinstance(config_digest, str) and DIGEST.fullmatch(config_digest), "Invalid runtime config digest")
        attestation = fetch(attestations[digest])
        statements = [verify_statement(fetch_blob(layer), layer, digest) for layer in attestation.get("layers", [])]
        predicates = {statement["predicateType"] for statement in statements}
        require(predicates & SLSA_TYPES, "No provenance predicate for " + platform)
        require(SPDX_TYPE in predicates, "No SPDX SBOM predicate for " + platform)
        results.append({"platform": platform, "manifestDigest": digest, "configDigest": config_digest,
                        "attestationDigest": attestations[digest], "predicates": sorted(predicates), "statements": statements})
    return results


class Registry:
    def __init__(self, repository):
        self.repository = repository
        self.metadata = urllib.request.build_opener(NoRedirect())
        self.blobs = urllib.request.build_opener(BlobRedirect())
        scope = urllib.parse.urlencode({"service": "ghcr.io", "scope": "repository:" + repository + ":pull"})
        self.token = json.loads(read(self.metadata, "https://ghcr.io/token?" + scope, LIMIT))["token"]
        require(isinstance(self.token, str) and self.token, "No public registry token")

    def request(self, kind, reference):
        return urllib.request.Request("https://ghcr.io/v2/" + self.repository + "/" + kind + "/" + reference,
                                      headers={"Authorization": "Bearer " + self.token, "Accept": ACCEPT})

    def fetch(self, digest):
        require(isinstance(digest, str) and DIGEST.fullmatch(digest), "Invalid manifest digest")
        raw = read(self.metadata, self.request("manifests", digest), LIMIT)
        require("sha256:" + hashlib.sha256(raw).hexdigest() == digest, "Manifest size/digest mismatch")
        return json.loads(raw)

    def fetch_blob(self, descriptor):
        name, _ = blob_descriptor(descriptor)
        raw = read(self.blobs, self.request("blobs", name), BLOB_LIMIT)
        return verify_blob(raw, descriptor)

    def verify_tag(self, tag, digest):
        require(re.fullmatch(r"[A-Za-z0-9_][A-Za-z0-9_.-]{0,127}", tag), "Invalid OCI tag")
        raw = read(self.metadata, self.request("manifests", tag), LIMIT)
        require("sha256:" + hashlib.sha256(raw).hexdigest() == digest, "Published tag points to a different index")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository", required=True)
    parser.add_argument("--digest", required=True)
    parser.add_argument("--source-manifest", type=Path, required=True)
    parser.add_argument("--tag", help="Also require this public tag to resolve to --digest")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    require(re.fullmatch(r"[a-z0-9._-]+/[a-z0-9._-]+", args.repository), "Expected lowercase GHCR owner/image")
    require(DIGEST.fullmatch(args.digest), "Expected sha256 index digest")
    registry = Registry(args.repository)
    if args.tag:
        registry.verify_tag(args.tag, args.digest)
    source = json.loads(args.source_manifest.read_text())
    results = verify(registry.fetch(args.digest), registry.fetch, source, registry.fetch_blob)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps({"indexDigest": args.digest, "sourceManifestSha256": hashlib.sha256(args.source_manifest.read_bytes()).hexdigest(), "images": results}, indent=2) + "\n")
    print("Verified platform manifest/source bindings, attestation content hashes, subjects and required predicate fields")


if __name__ == "__main__":
    try:
        main()
    except ValueError as error:
        # Transport errors are wrapped above; these messages contain no request URLs.
        raise SystemExit("Container release verification failed: " + str(error)) from None
    except (OSError, KeyError, TypeError):
        # Do not print exception chains containing signed blob URLs or raw metadata.
        raise SystemExit("Container release verification failed; inspect inputs and registry state") from None
