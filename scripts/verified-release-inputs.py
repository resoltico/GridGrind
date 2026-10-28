#!/usr/bin/env python3
"""Reuse completed release checks only for byte-identical staged images of the same tag commit."""
import argparse
import json
import os
from pathlib import Path
import re
import urllib.parse
import urllib.request

LIMIT = 4 * 1024 * 1024
REQUIRED = (
    "Check and build versioned delivery archives",
    "Build both runtime platforms without publishing",
    "Verify the exact staged images on both architectures",
    "Prepare and verify corresponding container source",
    "Validate retention access and incoming-release plan without deleting",
)
PUBLICATION = "Stage release assets and make sources publicly accessible"


def require(condition, message):
    if not condition:
        raise ValueError(message)


class LogRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        target = urllib.parse.urljoin(request.full_url, new_url)
        parsed = urllib.parse.urlsplit(target)
        require(parsed.scheme == "https" and parsed.hostname and not parsed.username and not parsed.password,
                "Release log redirect must use HTTPS without credentials")
        redirected = super().redirect_request(request, response, code, message, headers, target)
        if redirected is not None:
            for collection in (redirected.headers, redirected.unredirected_hdrs):
                for name in list(collection):
                    if name.lower() in {"authorization", "proxy-authorization", "cookie"}:
                        del collection[name]
        return redirected


class Api:
    def __init__(self, repository, token):
        self.base = "https://api.github.com/repos/" + repository
        self.token = token

    def get(self, route, log=False):
        request = urllib.request.Request(self.base + route, headers={
            "Authorization": "Bearer " + self.token,
            "Accept": "application/vnd.github+json",
            "User-Agent": "GridGrind-verified-release-inputs",
        })
        opener = urllib.request.build_opener(LogRedirect() if log else NoRedirect())
        with opener.open(request, timeout=30) as response:
            raw = response.read(LIMIT + 1)
        require(len(raw) <= LIMIT, "Release verification evidence exceeds the size bound")
        return raw.decode("utf-8-sig") if log else json.loads(raw)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *_):
        raise ValueError("Release evidence API metadata must not redirect")


def verified_images(api, repository, run_id, version, commit):
    run = api.get("/actions/runs/" + run_id)
    require(run.get("id") == int(run_id) and type(run.get("run_attempt")) is int
            and run["run_attempt"] > 0 and run.get("repository", {}).get("full_name") == repository
            and run.get("path") == ".github/workflows/release.yml"
            and run.get("head_sha") == commit and run.get("head_branch") == "v" + version
            and run.get("event") == "push" and run.get("status") == "completed"
            and run.get("conclusion") == "failure", "Prior run is not the completed release attempt for this tag commit")
    jobs = api.get("/actions/runs/" + run_id + "/jobs?per_page=100")
    require(jobs.get("total_count") == 1 and len(jobs.get("jobs", [])) == 1,
            "Expected exactly one prior release job")
    job = jobs["jobs"][0]
    require(job.get("run_id") == int(run_id) and job.get("run_attempt") == run["run_attempt"]
            and job.get("name") == "release" and job.get("status") == "completed"
            and job.get("conclusion") == "failure", "Prior release job is not a completed failure")
    steps = job.get("steps", [])
    for name in REQUIRED:
        matches = [step for step in steps if step.get("name") == name]
        require(len(matches) == 1 and matches[0].get("status") == "completed"
                and matches[0].get("conclusion") == "success", "Prior verification did not pass: " + name)
    require([step.get("name") for step in steps if step.get("conclusion") == "failure"] == [PUBLICATION],
            "Prior run did not stop solely at initial release publication")
    require(type(job.get("id")) is int and job["id"] > 0, "Invalid prior job identity")
    log = api.get("/actions/jobs/" + str(job["id"]) + "/logs", log=True)
    log = re.sub(r"\x1b\[[0-9;]*m", "", log)
    blocks = re.findall(r"^\S+Z ##\[group\]Run \./gradlew containerTest [^\n]*\n(.*?)^\S+Z ##\[endgroup\]\s*$",
                        log, re.MULTILINE | re.DOTALL)
    require(len(blocks) == 1, "Expected exactly one prior container verification command block")
    command = (r'^\S+Z \./gradlew containerTest "-Pversion=' + re.escape(version)
               + r'" "-PcontainerImage=(sha256:[0-9a-f]{64})"\s*$')
    images = re.findall(command, blocks[0], re.MULTILINE)
    require(len(images) == 2 and len(set(images)) == 2,
            "Prior product verification must identify exactly two distinct staged images")
    return {"amd64": images[0], "arm64": images[1]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("repo", "run", "version", "commit", "output"):
        parser.add_argument("--" + name, required=True)
    args = parser.parse_args()
    require(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", args.repo), "Invalid repository")
    require(re.fullmatch(r"[1-9][0-9]*", args.run), "Invalid prior run ID")
    require(re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", args.version), "Invalid release version")
    require(re.fullmatch(r"[0-9a-f]{40}", args.commit), "Invalid tag commit")
    images = verified_images(Api(args.repo, os.environ["GH_TOKEN"]), args.repo, args.run, args.version, args.commit)
    Path(args.output).write_text(json.dumps({"verifiedRun": args.run, "commit": args.commit, "images": images}, indent=2) + "\n")
    for architecture, image in images.items():
        print(architecture + "=" + image)


if __name__ == "__main__":
    main()
