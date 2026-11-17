#!/usr/bin/env python3
"""Stage verified release assets, expose sources as a prerelease, then finalize.

The workflow must serialize publication and verify images before finalize. API
mutations and uploads are never automatically retried after an ambiguous failure.
"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
import http.client
import urllib.error
import urllib.request
from dataclasses import dataclass
from release_confirmation import ReleaseFailure, confirm, diagnostic, tool_write


class PublicationError(ReleaseFailure):
    pass


class ApiError(PublicationError):
    def __init__(self, status, method="GET", request_id=None):
        self.status = status
        super().__init__("GitHub API returned HTTP " + str(status),
                         uncertain=method != "GET" and (status in (408, 429) or status >= 500),
                         category="http", status=status, request_id=request_id)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        raise PublicationError("Unexpected GitHub API redirect")


class Github:
    def __init__(self, repo, token):
        self.base = "https://api.github.com/repos/" + repo
        self.token = token
        self.opener = urllib.request.build_opener(NoRedirect())

    def request(self, method, route, body=None):
        if not route.startswith("/") or route.startswith("//") or "#" in route:
            raise PublicationError("Invalid GitHub API route")
        headers = {
            "Authorization": "Bearer " + self.token,
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2026-03-10",
            "User-Agent": "GridGrind-release-publication",
        }
        payload = None if body is None else json.dumps(body).encode("utf-8")
        if payload is not None:
            headers["Content-Type"] = "application/json"
        attempts = 3 if method == "GET" else 1
        for attempt in range(attempts):
            try:
                request = urllib.request.Request(self.base + route, data=payload, headers=headers, method=method)
                with self.opener.open(request, timeout=20) as response:
                    raw = response.read(4 * 1024 * 1024 + 1)
                if len(raw) > 4 * 1024 * 1024:
                    raise PublicationError("GitHub API response exceeds the size bound")
                return json.loads(raw) if raw else None
            except urllib.error.HTTPError as error:
                status = error.code
                request_id = error.headers.get("X-GitHub-Request-Id") if error.headers else None
                error.close()
                retryable = status == 429 or 500 <= status < 600
                if method != "GET" or not retryable or attempt + 1 == attempts:
                    raise ApiError(status, method, request_id) from None
            except (urllib.error.URLError, TimeoutError, OSError, http.client.HTTPException) as error:
                if attempt + 1 == attempts:
                    raise PublicationError("GitHub API request failed or timed out; inspect remote state before retrying", uncertain=method != "GET", category="transport") from None
            except (ValueError, UnicodeError) as error:
                raise PublicationError("GitHub API returned invalid JSON", uncertain=method != "GET", category="response") from None
            time.sleep(attempt + 1)
        raise PublicationError("GitHub API request did not complete")


@dataclass(frozen=True)
class Asset:
    path: Path
    name: str
    size: int
    digest: str


def local_assets(paths):
    result = []
    for value in paths:
        path = Path(value).resolve()
        if not path.is_file() or "#" in str(path) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", path.name):
            raise PublicationError("Expected a regular asset with an unambiguous release filename")
        checksum = hashlib.sha256()
        size = 0
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                size += len(block)
                checksum.update(block)
        result.append(Asset(path, path.name, size, "sha256:" + checksum.hexdigest()))
    if not result or len({asset.name for asset in result}) != len(result):
        raise PublicationError("Expected nonempty assets with distinct filenames")
    return tuple(result)


def upload(repo, tag, assets):
    environment = dict(os.environ)
    environment["GH_HOST"] = "github.com"
    environment.pop("GH_DEBUG", None)
    command = ["gh", "release", "upload", tag, "--repo", repo, "--clobber"]
    command.extend(str(asset.path) for asset in assets)
    tool_write(command, env=environment, capture_output=True, timeout=600)


def validate_release_identity(release):
    if (not isinstance(release, dict) or type(release.get("id")) is not int
            or release["id"] <= 0 or not isinstance(release.get("tag_name"), str)
            or not release["tag_name"] or type(release.get("draft")) is not bool
            or type(release.get("prerelease")) is not bool):
        raise PublicationError("GitHub release identity or state is invalid")
    return release


class Publication:
    def __init__(self, api, repo, version, commit, notes, assets, uploader=upload, recorder=None):
        self.api = api
        self.repo = repo
        self.version = version
        self.tag = "v" + version
        self.commit = commit
        self.notes = notes
        self.assets = assets
        self.uploader = uploader
        self.observations = []
        self.recorder = recorder or self.observations.append

    def validate_tag(self):
        record = self.api.request("GET", "/git/ref/tags/" + self.tag)
        seen = set()
        for _ in range(16):
            target = record.get("object") if isinstance(record, dict) else None
            if (not isinstance(target, dict) or not isinstance(target.get("sha"), str)
                    or not re.fullmatch(r"[0-9a-f]{40}", target["sha"])):
                raise PublicationError("GitHub tag has an invalid object identity")
            sha = target["sha"]
            if target.get("type") == "commit":
                if sha != self.commit:
                    raise PublicationError("Release tag does not resolve to the verified commit")
                return
            if target.get("type") != "tag" or sha in seen:
                raise PublicationError("Invalid annotated tag chain")
            seen.add(sha)
            record = self.api.request("GET", "/git/tags/" + sha)
        raise PublicationError("Annotated tag chain exceeds the bound")

    def release(self):
        try:
            release = validate_release_identity(self.api.request("GET", "/releases/tags/" + self.tag))
        except ApiError as error:
            if error.status != 404:
                raise
            # GitHub's tag endpoint does not expose drafts. An authenticated
            # release-list lookup must rule those out before creating a release.
            return self.listed_release()
        if release["tag_name"] != self.tag:
            raise PublicationError("GitHub release tag does not match the requested identity")
        return release

    def listed_release(self):
        matches = []
        seen = set()
        for page in range(1, 101):
            rows = self.api.request("GET", "/releases?per_page=100&page=" + str(page))
            if not isinstance(rows, list) or len(rows) > 100:
                raise PublicationError("GitHub release list page is invalid")
            for row in rows:
                release = validate_release_identity(row)
                if release["id"] in seen:
                    raise PublicationError("GitHub release pagination returned duplicate IDs")
                seen.add(release["id"])
                if release["tag_name"] == self.tag:
                    matches.append(release)
                    if len(matches) > 1:
                        raise PublicationError("Multiple GitHub releases match the requested tag")
            if len(rows) < 100:
                return matches[0] if matches else None
        raise PublicationError("GitHub release list pagination exceeds the bound")

    def remote_assets(self, release):
        result = []
        for page in range(1, 101):
            rows = self.api.request("GET", "/releases/%s/assets?per_page=100&page=%s" % (release["id"], page))
            if not isinstance(rows, list) or len(rows) > 100 or any(not isinstance(row, dict) for row in rows):
                raise PublicationError("GitHub release asset page is invalid")
            if any(not isinstance(row.get("name"), str) or not row["name"] for row in rows):
                raise PublicationError("GitHub release asset name is invalid")
            result.extend(rows)
            if len(rows) < 100:
                return result
        raise PublicationError("GitHub release asset pagination exceeds the bound")

    def verify(self, release, state):
        if release is None:
            raise PublicationError("Release is absent; prepare it first")
        if release.get("tag_name") != self.tag or release.get("target_commitish") != self.commit:
            raise PublicationError("Release identity differs from the verified commit")
        actual_state = "draft" if release["draft"] else "prerelease" if release["prerelease"] else "stable"
        if actual_state != state:
            raise PublicationError("Expected %s release, found %s" % (state, actual_state))
        if release.get("body") != self.notes:
            raise PublicationError("Release notes do not match the verified local notes")
        remote = self.remote_assets(release)
        names = [row.get("name") for row in remote]
        if len(set(names)) != len(names) or set(names) != {asset.name for asset in self.assets}:
            raise PublicationError("Release asset names do not exactly match verified local assets")
        by_name = {row["name"]: row for row in remote}
        for asset in self.assets:
            row = by_name[asset.name]
            if (row.get("digest") != asset.digest or type(row.get("size")) is not int
                    or row["size"] != asset.size or row.get("state") != "uploaded"):
                raise PublicationError("Release asset digest, size or upload state mismatch: " + asset.name)

    def require_draft(self, identifier):
        current = self.release()
        if current is None or current["id"] != identifier or not current["draft"]:
            raise PublicationError("Release is no longer the expected draft; no public assets will be overwritten")
        if current.get("target_commitish") != self.commit:
            raise PublicationError("Draft target differs from the verified commit")
        return current

    def release_fields(self, draft, prerelease, latest):
        # Send the complete intended release identity/state on every mutation.
        # GitHub's observed draft PATCH semantics can reset an omitted tag name.
        return {
            "tag_name": self.tag, "target_commitish": self.commit,
            "name": "GridGrind " + self.version, "body": self.notes,
            "draft": draft, "prerelease": prerelease,
            "make_latest": "true" if latest else "false",
        }

    def observe_release(self, identifier, state, prior=None, assets=True, allow_absent=False):
        current = self.release()
        if current is None:
            if allow_absent:
                return None
            raise PublicationError("Previously observed release disappeared")
        if (current["tag_name"] != self.tag or current.get("target_commitish") != self.commit
                or identifier is not None and current["id"] != identifier):
            raise PublicationError("Release identity differs from the intended publication")
        actual = "draft" if current["draft"] else "prerelease" if current["prerelease"] else "stable"
        relevant = lambda row: {k: row.get(k) for k in ("id", "tag_name", "target_commitish", "body", "draft", "prerelease", "name")}
        if actual != state or current.get("body") != self.notes or current.get("name") != "GridGrind " + self.version:
            if prior is not None and relevant(current) == relevant(prior):
                if assets:
                    self.verify(current, actual)
                return None
            raise PublicationError("Release state or notes contradict the intended publication")
        if assets:
            self.verify(current, state)
        return current

    def transition(self, operation, release, state, fields):
        return confirm(operation,
                       lambda: self.api.request("PATCH", "/releases/" + str(release["id"]), fields),
                       lambda: self.observe_release(release["id"], state, release), self.recorder)

    def run(self, stage):
        self.validate_tag()
        release = self.release()
        if release is not None and not release["draft"] and not release["prerelease"]:
            raise PublicationError("Stable public release exists; use read-only manual postflight")
        if stage == "prepare":
            if release is not None and not release["draft"]:
                self.verify(release, "prerelease")
                return
            if release is None:
                acknowledged = []
                def create():
                    row = self.api.request("POST", "/releases", self.release_fields(True, True, False))
                    if row is not None:
                        acknowledged.append(validate_release_identity(row)["id"])
                release = confirm("release-create", create,
                                  lambda: self.observe_release(acknowledged[0] if acknowledged else None,
                                                               "draft", assets=False, allow_absent=True), self.recorder)
            current = self.require_draft(release["id"])
            before = self.remote_assets(current)
            names = [row.get("name") for row in before]
            if len(set(names)) != len(names) or set(names) - {asset.name for asset in self.assets}:
                raise PublicationError("Draft has unexpected assets; review them before resuming")
            self.require_draft(release["id"])
            # This update deliberately replaces the exact preflight draft metadata.
            confirm("release-draft-update",
                    lambda: self.api.request("PATCH", "/releases/" + str(release["id"]), self.release_fields(True, True, False)),
                    lambda: self.observe_release(release["id"], "draft", current, assets=False), self.recorder)
            self.require_draft(release["id"])
            def uploaded():
                observed = self.observe_release(release["id"], "draft", assets=False)
                if observed is None:
                    return None
                rows = self.remote_assets(observed)
                by_name = {row["name"]: row for row in rows}
                if len(by_name) != len(rows) or set(by_name) - {asset.name for asset in self.assets}:
                    raise PublicationError("Draft has unexpected or duplicate assets")
                expected = {asset.name: asset for asset in self.assets}
                prior = {row["name"]: row for row in before}
                pending = len(rows) != len(expected)
                for name, row in by_name.items():
                    asset = expected[name]
                    if (type(row.get("size")) is not int
                            or (row.get("digest"), row.get("size"), row.get("state")) != (asset.digest, asset.size, "uploaded")):
                        if row != prior.get(name):
                            raise PublicationError("Draft asset differs from both reviewed bytes and exact prior state")
                        pending = True
                return None if pending else observed
            confirm("release-assets-upload", lambda: self.uploader(self.repo, self.tag, self.assets), uploaded, self.recorder)
        elif stage == "expose":
            if release is None:
                raise PublicationError("Cannot expose an absent release")
            state = "draft" if release["draft"] else "prerelease"
            self.verify(release, state)
            if state == "draft":
                self.require_draft(release["id"])
                self.transition("release-expose", release, "prerelease", self.release_fields(False, True, False))
        elif stage == "finalize":
            self.verify(release, "prerelease")
            try:
                prior_latest = self.api.request("GET", "/releases/latest")
            except ApiError as error:
                if error.status != 404:
                    raise
                prior_latest = None
            self.transition("release-finalize", release, "stable", self.release_fields(False, False, True))
            def latest():
                try:
                    row = self.api.request("GET", "/releases/latest")
                except ApiError as error:
                    if error.status == 404 and prior_latest is None:
                        return None
                    raise
                if isinstance(row, dict) and row.get("id") == release["id"] and row.get("tag_name") == self.tag:
                    self.verify(row, "stable")
                    return self.observe_release(release["id"], "stable")
                if row == prior_latest:
                    return None
                raise PublicationError("GitHub latest points to an unexpected release")
            # The PATCH already happened; this operation performs observations only.
            confirm("github-latest-readback", lambda: None, latest, self.recorder)
        else:
            raise PublicationError("Unknown publication stage")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", choices=("prepare", "expose", "finalize"))
    parser.add_argument("--repo", required=True)
    parser.add_argument("--version", required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--notes", required=True)
    parser.add_argument("--asset", action="append", required=True)
    args = parser.parse_args(argv)
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", args.repo):
        raise PublicationError("Expected GitHub owner/repository")
    if not re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", args.version):
        raise PublicationError("Expected release version X.Y.Z")
    if not re.fullmatch(r"[0-9a-f]{40}", args.commit):
        raise PublicationError("Expected the verified commit SHA")
    notes_path = Path(args.notes)
    if notes_path.stat().st_size > 1024 * 1024:
        raise PublicationError("Release notes exceed the size bound")
    notes = notes_path.read_text(encoding="utf-8")
    if not notes.strip():
        raise PublicationError("Release notes must not be empty")
    assets = local_assets(args.asset)
    token = os.environ.get("GH_TOKEN")
    if not token:
        raise PublicationError("GH_TOKEN is required")
    publication = Publication(Github(args.repo, token), args.repo, args.version, args.commit, notes, assets)
    try:
        publication.run(args.stage)
    finally:
        for observation in publication.observations:
            print(json.dumps(observation, sort_keys=True), file=sys.stderr)
    print("Release publication stage verified: " + args.stage)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (ReleaseFailure, OSError, UnicodeError) as error:
        print("Release publication stopped: " + str(diagnostic(error)), file=sys.stderr)
        sys.exit(1)
