#!/usr/bin/env python3
"""Plan GHCR retention from verified manifest graphs; apply only with --apply.

Keep the five newest stable-tagged roots and every other nonretired tagged root.
Unknown untagged records are not garbage-collected. Publication must be serialized
with cleanup: GitHub provides no atomic snapshot-and-delete package operation.
"""

import argparse
import datetime
import hashlib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass

DIGEST = re.compile(r"sha256:[0-9a-f]{64}\Z")
STABLE = re.compile(r"[0-9]+\.[0-9]+\.[0-9]+\Z")
SOURCES = re.compile(r"([0-9]+\.[0-9]+\.[0-9]+)-sources\Z")
INDEX_TYPES = {
    "application/vnd.oci.image.index.v1+json",
    "application/vnd.docker.distribution.manifest.list.v2+json",
}
MANIFEST_TYPES = {
    "application/vnd.oci.image.manifest.v1+json",
    "application/vnd.docker.distribution.manifest.v2+json",
}
ACCEPT = ", ".join(sorted(INDEX_TYPES | MANIFEST_TYPES))
MAX_BODY = 4 * 1024 * 1024
MAX_PAGES = 100
MAX_GRAPH_NODES = 10000


class PruningError(Exception):
    """Unproven state must never become a deletion plan."""


class HttpFailure(PruningError):
    def __init__(self, method, status):
        self.status = status
        super().__init__("HTTP %s failed with status %s" % (method, status))


class MissingManifest(PruningError):
    """An explicit registry 404; only retired graph traversal may tolerate it."""


def digest(value):
    if not isinstance(value, str) or not DIGEST.fullmatch(value):
        raise PruningError("Expected a lowercase sha256 manifest digest")
    return value


def parse_json(data):
    try:
        return json.loads(data)
    except (ValueError, UnicodeError) as error:
        raise PruningError("Invalid JSON response") from error


def utc_timestamp(value):
    if not isinstance(value, str):
        raise PruningError("Package version lacks a creation timestamp")
    try:
        result = datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise PruningError("Invalid package creation timestamp") from error
    if result.tzinfo is None:
        raise PruningError("Package creation timestamp lacks a timezone")
    return result.astimezone(datetime.timezone.utc)


@dataclass(frozen=True)
class Version:
    identifier: int
    name: str
    tags: tuple
    created_at: str

    @staticmethod
    def parse(row):
        if not isinstance(row, dict):
            raise PruningError("Package API version must be an object")
        identifier = row.get("id")
        if type(identifier) is not int or identifier <= 0:
            raise PruningError("Invalid package version ID")
        try:
            tags = row["metadata"]["container"]["tags"]
        except (KeyError, TypeError) as error:
            raise PruningError("Package version lacks container tags") from error
        if not isinstance(tags, list) or any(not isinstance(tag, str) or not tag for tag in tags):
            raise PruningError("Invalid container tags")
        if len(set(tags)) != len(tags):
            raise PruningError("Duplicate container tags")
        created_at = row.get("created_at")
        utc_timestamp(created_at)
        return Version(identifier, digest(row.get("name")), tuple(sorted(tags)), created_at)


def snapshot(versions):
    identifiers = [version.identifier for version in versions]
    if len(set(identifiers)) != len(identifiers):
        raise PruningError("Package pagination returned duplicate IDs; retry from a stable state")
    if len({version.name for version in versions}) != len(versions):
        raise PruningError("Package pagination returned duplicate manifest names")
    tag_owners = {}
    for version in versions:
        for tag in version.tags:
            if tag in tag_owners:
                raise PruningError("Package pagination returned duplicate tag ownership")
            tag_owners[tag] = version.identifier
    return tuple(sorted(versions, key=lambda version: version.identifier))


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        # Never forward a GitHub or registry credential to a redirected host.
        raise PruningError("Unexpected HTTP redirect; no deletion is safe")


class Http:
    def __init__(self):
        self.opener = urllib.request.build_opener(NoRedirect())

    def request(self, method, url, headers=None):
        attempts = 3 if method == "GET" else 1
        for attempt in range(attempts):
            request = urllib.request.Request(url, method=method, headers=headers or {})
            try:
                with self.opener.open(request, timeout=15) as response:
                    body = response.read(MAX_BODY + 1)
                    if len(body) > MAX_BODY:
                        raise PruningError("HTTP response exceeds the bounded body size")
                    return body
            except urllib.error.HTTPError as error:
                retryable = error.code == 429 or 500 <= error.code < 600
                error.close()
                if not retryable or attempt + 1 == attempts:
                    raise HttpFailure(method, error.code) from error
            except (urllib.error.URLError, TimeoutError, OSError) as error:
                if attempt + 1 == attempts:
                    raise PruningError("HTTP %s failed or timed out" % method) from error
            time.sleep(attempt + 1)
        raise PruningError("HTTP request did not complete")


class Github:
    def __init__(self, http, owner, package, token):
        self.http = http
        self.headers = {
            "Authorization": "Bearer " + token,
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2026-03-10",
            "User-Agent": "GridGrind-container-retention",
        }
        result = parse_json(self.http.request("GET", "https://api.github.com/users/" + owner, self.headers))
        owner_type = result.get("type") if isinstance(result, dict) else None
        if owner_type not in ("User", "Organization"):
            raise PruningError("Unsupported GitHub package owner type")
        route = "orgs" if owner_type == "Organization" else "users"
        self.endpoint = "https://api.github.com/%s/%s/packages/container/%s/versions" % (route, owner, package)

    def versions(self):
        versions = []
        for page in range(1, MAX_PAGES + 1):
            url = self.endpoint + "?per_page=100&page=" + str(page)
            rows = parse_json(self.http.request("GET", url, self.headers))
            if not isinstance(rows, list) or len(rows) > 100:
                raise PruningError("Invalid package version page")
            versions.extend(Version.parse(row) for row in rows)
            if len(rows) < 100:
                return snapshot(versions)
        raise PruningError("Package pagination limit reached; no deletion is safe")

    def delete(self, version):
        self.http.request("DELETE", self.endpoint + "/" + str(version.identifier), self.headers)


class Registry:
    def __init__(self, http, owner, package):
        self.http = http
        self.repository = owner.lower() + "/" + package.lower()
        self.token = None

    def manifest(self, requested_digest):
        digest(requested_digest)
        if self.token is None:
            query = urllib.parse.urlencode({"service": "ghcr.io", "scope": "repository:" + self.repository + ":pull"})
            result = parse_json(self.http.request("GET", "https://ghcr.io/token?" + query))
            self.token = result.get("token") if isinstance(result, dict) else None
            if not isinstance(self.token, str) or not self.token:
                raise PruningError("Public GHCR pull token was not provided")
        url = "https://ghcr.io/v2/" + self.repository + "/manifests/" + requested_digest
        try:
            return self.http.request("GET", url, {"Authorization": "Bearer " + self.token, "Accept": ACCEPT})
        except HttpFailure as error:
            if error.status == 404:
                raise MissingManifest("Registry manifest is unavailable: " + requested_digest) from error
            raise


class ManifestGraph:
    def __init__(self, registry):
        self.registry = registry
        self.edges = {}
        self.missing = set()

    def children(self, requested_digest):
        digest(requested_digest)
        if requested_digest in self.edges:
            return self.edges[requested_digest]
        if requested_digest in self.missing:
            raise MissingManifest("Registry manifest is unavailable: " + requested_digest)
        if len(self.edges) >= MAX_GRAPH_NODES:
            raise PruningError("Manifest graph limit reached")
        try:
            raw = self.registry.manifest(requested_digest)
        except MissingManifest:
            self.missing.add(requested_digest)
            raise
        actual_digest = "sha256:" + hashlib.sha256(raw).hexdigest()
        if actual_digest != requested_digest:
            raise PruningError("Registry manifest bytes do not match requested digest")
        document = parse_json(raw)
        if not isinstance(document, dict) or document.get("schemaVersion") != 2:
            raise PruningError("Unsupported registry manifest schema")
        media_type = document.get("mediaType")
        children = []
        if media_type in INDEX_TYPES:
            descriptors = document.get("manifests")
            if not isinstance(descriptors, list):
                raise PruningError("OCI index lacks manifest descriptors")
            for descriptor in descriptors:
                if not isinstance(descriptor, dict):
                    raise PruningError("Invalid OCI manifest descriptor")
                children.append(digest(descriptor.get("digest")))
        elif media_type not in MANIFEST_TYPES:
            raise PruningError("Unsupported registry manifest media type")
        # Config/layers are blobs, not manifest package versions. An OCI subject
        # is another manifest and must be protected together with its referrer.
        if "subject" in document:
            subject = document["subject"]
            if not isinstance(subject, dict):
                raise PruningError("Invalid OCI subject descriptor")
            children.append(digest(subject.get("digest")))
        self.edges[requested_digest] = tuple(children)
        return self.edges[requested_digest]

    def reachable(self, roots, allow_missing=False):
        visited = set()
        resolved = set()
        pending = list(roots)
        while pending:
            current = pending.pop()
            if current in visited:
                continue
            visited.add(current)
            if len(visited) > MAX_GRAPH_NODES:
                raise PruningError("Manifest graph limit reached")
            try:
                children = self.children(current)
            except MissingManifest:
                if not allow_missing:
                    raise
                # A retired tombstone proves no bytes or descendants. Never
                # nominate its package ID or infer any unknown orphan children.
                continue
            resolved.add(current)
            pending.extend(children)
        return resolved


@dataclass(frozen=True)
class Plan:
    initial_state: tuple
    retained: tuple
    retired: tuple
    deletions: tuple
    protected_digest_count: int
    unavailable_retired: tuple = ()
    retired_sources: tuple = ()

    def report(self):
        def describe(version):
            return {"id": version.identifier, "digest": version.name, "tags": version.tags}
        return {
            "retained_stable_roots": [describe(version) for version in self.retained],
            "retired_stable_roots": [describe(version) for version in self.retired],
            "retired_source_roots": [describe(version) for version in self.retired_sources],
            "protected_manifest_count": self.protected_digest_count,
            "unavailable_retired_manifests": self.unavailable_retired,
            "delete": [describe(version) for version in self.deletions],
        }


def make_plan(versions, graph, keep=5):
    if keep <= 0:
        raise PruningError("Retention count must be positive")
    initial = snapshot(versions)
    stable = sorted(
        (version for version in initial if any(STABLE.fullmatch(tag) for tag in version.tags)),
        key=lambda version: (utc_timestamp(version.created_at), version.identifier), reverse=True,
    )
    retained, retired = tuple(stable[:keep]), tuple(stable[keep:])
    if not retired:
        return Plan(initial, retained, retired, (), 0)
    retired_ids = {version.identifier for version in retired}
    stable_owners = {tag: version.identifier for version in stable for tag in version.tags if STABLE.fullmatch(tag)}

    def retired_tag(tag):
        source = SOURCES.fullmatch(tag)
        base_tag = source.group(1) if source else tag
        # A source image published before its binary anchor remains protected.
        return stable_owners.get(base_tag) in retired_ids

    retired_sources = tuple(version for version in initial if any(
        SOURCES.fullmatch(tag) and retired_tag(tag) for tag in version.tags
    ))
    # A retained source tag or any unrelated alias independently protects its
    # graph, even when the same root also carries a retired release/source tag.
    protected_roots = [version.name for version in initial if any(not retired_tag(tag) for tag in version.tags)]
    protected = graph.reachable(protected_roots)
    retired_graph = graph.reachable(
        (version.name for version in retired + retired_sources), allow_missing=True,
    )
    eligible = retired_graph - protected
    deletions = tuple(version for version in initial if version.name in eligible)
    return Plan(initial, retained, retired, deletions, len(protected), tuple(sorted(graph.missing)), retired_sources)


def apply_plan(api, plan):
    # Validate the complete fresh state before the first destructive request.
    # Immutable manifest digests make a second graph download unnecessary.
    if snapshot(api.versions()) != plan.initial_state:
        raise PruningError("Package state changed after planning; no versions deleted")
    for version in plan.deletions:
        api.delete(version)


def preflight_keep(versions, keep, pending_version):
    if not pending_version:
        return keep
    if not STABLE.fullmatch(pending_version) or keep < 2:
        raise PruningError("Pending-release planning needs X.Y.Z and at least two retention slots")
    return keep if any(pending_version in version.tags for version in versions) else keep - 1


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--owner", required=True)
    parser.add_argument("--package", default="gridgrind")
    parser.add_argument("--keep", type=int, default=5)
    parser.add_argument("--apply", action="store_true", help="delete the fully verified plan; default is dry-run")
    parser.add_argument("--pending-version", help="dry-run only: reserve one slot if this runtime version is not published yet")
    args = parser.parse_args(argv)
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9-]{0,38}", args.owner):
        raise PruningError("Invalid GitHub owner")
    if not re.fullmatch(r"[a-z0-9][a-z0-9._-]*", args.package):
        raise PruningError("Expected a simple lowercase GHCR package name")
    token = os.environ.get("GH_TOKEN")
    if not token:
        raise PruningError("GH_TOKEN is required for the GitHub package API")
    if args.pending_version and args.apply:
        raise PruningError("Pending-release planning cannot delete versions")
    http = Http()
    api = Github(http, args.owner, args.package, token)
    versions = api.versions()
    effective_keep = preflight_keep(versions, args.keep, args.pending_version)
    plan = make_plan(versions, ManifestGraph(Registry(http, args.owner, args.package)), effective_keep)
    print(json.dumps({"mode": "apply" if args.apply else "dry-run", **plan.report()}, indent=2))
    if args.apply:
        apply_plan(api, plan)
        print("Deleted %s proven retired package versions" % len(plan.deletions))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PruningError as error:
        print("Container retention stopped: " + str(error), file=sys.stderr)
        sys.exit(1)
