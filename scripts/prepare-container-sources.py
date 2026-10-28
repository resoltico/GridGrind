#!/usr/bin/env python3
"""Prepare a release source archive for two locally built Debian/Corretto images."""

import argparse
import gzip
import hashlib
import json
import os
from pathlib import Path
import re
import shlex
import shutil
import signal
import subprocess
import sys
import tarfile
import time
from urllib.parse import urlparse
import uuid


ASSET_LIMIT = 2 * 1024**3 - 1
INDEX_LIMIT = 256 * 1024**2
PLATFORMS = {"linux/amd64": "amd64", "linux/arm64": "arm64"}
SOURCE_NAME = re.compile(r"[a-z0-9][a-z0-9+.-]*\Z")
SOURCE_VERSION = re.compile(r"[0-9][A-Za-z0-9.+:~\-]*\Z")
SHA256 = re.compile(r"[a-f0-9]{64}\Z")
MANIFEST_MEDIA_TYPES = {
    "application/vnd.oci.image.manifest.v1+json",
    "application/vnd.docker.distribution.manifest.v2+json",
}
BUILD_INPUTS = (
    "Dockerfile", "release/container-runtime.properties", "scripts/prepare-container-sources.py",
    "LICENSE", "NOTICE",
)


def require(condition, message):
    if not condition:
        raise ValueError(message)


def digest(path):
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def write_archive(content, archive, version):
    """Normalize only the outer delivery metadata; upstream archive contents stay unchanged."""
    def normalize(info):
        require(info.isdir() or info.isfile(), "Source delivery tree must contain only files and directories")
        info.uid = info.gid = 0
        info.uname = info.gname = "root"
        info.mtime = 0
        info.mode = 0o755 if info.isdir() else 0o644
        info.pax_headers = {}
        return info

    prefix = f"gridgrind-container-sources-{version}"
    with archive.open("wb") as target:
        with gzip.GzipFile(filename="", mode="wb", compresslevel=1, mtime=0, fileobj=target) as compressed:
            with tarfile.open(fileobj=compressed, mode="w") as tar:
                tar.add(content, arcname=prefix, recursive=False, filter=normalize)
                for path in sorted(content.rglob("*"), key=lambda p: p.relative_to(content).as_posix()):
                    tar.add(path, arcname=prefix + "/" + path.relative_to(content).as_posix(),
                            recursive=False, filter=normalize)


def copy_build_inputs(repository, content):
    records = []
    for name in BUILD_INPUTS:
        source = repository / name
        require(source.is_file(), "Missing container build source: " + name)
        target = content / "gridgrind-build" / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
        records.append({"sourcePath": name, "archivePath": target.relative_to(content).as_posix(),
                        "sha256": digest(target), "size": target.stat().st_size})
    return records


def controls(text):
    """Read Debian control stanzas, including a clear-signed .dsc body."""
    if text.startswith("-----BEGIN PGP SIGNED MESSAGE-----"):
        text = text.split("\n\n", 1)[1].split("-----BEGIN PGP SIGNATURE-----", 1)[0]
        text = "\n".join(line[2:] if line.startswith("- ") else line for line in text.splitlines())
    result = []
    for paragraph in re.split(r"\n\s*\n", text.strip()):
        fields = {}
        current = None
        for line in paragraph.splitlines():
            if line.startswith((" ", "\t")):
                require(current is not None, "Control continuation without a field")
                fields[current] += "\n" + line[1:]
            else:
                key, separator, value = line.partition(":")
                require(separator and key not in fields, "Malformed or duplicated control field")
                current = key
                fields[key] = value.strip()
        result.append(fields)
    return result


def checksums(value):
    files = {}
    for line in value.strip().splitlines():
        parts = line.split()
        require(len(parts) == 3, "Malformed Checksums-Sha256 row")
        sha, size, name = parts
        require(SHA256.fullmatch(sha), "Invalid source SHA256")
        require(Path(name).name == name and name not in {".", ".."}, "Unsafe source filename")
        require(name not in files and size.isdecimal(), "Duplicate file or invalid source size")
        files[name] = {"sha256": sha, "size": int(size)}
    require(files, "Empty source checksums")
    return files


def source_record(text, name, version):
    matches = [r for r in controls(text) if r.get("Package") == name and r.get("Version") == version]
    # Multiple archive snapshots can carry the same immutable source stanza.
    require(matches, f"No authenticated source metadata for {name}={version}")
    files = checksums(matches[0].get("Checksums-Sha256", ""))
    require(all(checksums(r.get("Checksums-Sha256", "")) == files for r in matches),
            f"Conflicting source archives for {name}={version}")
    require(sum(n.endswith(".dsc") for n in files) == 1, "Source metadata needs exactly one .dsc")
    return files


def require_debian_uri(uri):
    parsed = urlparse(uri)
    host = parsed.hostname or ""
    require(parsed.scheme in {"http", "https"} and (host == "debian.org" or host.endswith(".debian.org"))
            and parsed.username is None and parsed.password is None and not parsed.fragment,
            f"Unsupported or non-Debian source archive URI: {uri}")


def source_config(text, filename):
    """Preserve archive/keyring configuration; enable source indexes, including retained snapshots."""
    require(not re.search(r"(?im)(?:trusted|allow-insecure|allow-unauthenticated|allow-weak|allow-downgrade-to-insecure)"
                          r"\s*[:=]\s*(?:yes|true|1)\b", text),
            "Unauthenticated apt source configuration is not accepted")
    if filename.endswith(".list"):
        lines = []
        for line in text.splitlines():
            if re.match(r"^\s*deb(?:-src)?\s", line):
                lines.append(re.sub(r"^(\s*)deb(?:-src)?\s", r"\1deb-src ", line))
        converted = "\n".join(lines) + "\n"
        uris = re.findall(r"(?m)^\s*deb-src\s+(?:\[[^\]]*\]\s+)?(\S+)", converted)
    else:
        blocks = []
        for block in re.split(r"\n\s*\n", text.strip()):
            if re.search(r"(?im)^Enabled:\s*no\s*$", block):
                continue
            if not re.search(r"(?m)^Types:\s*.*\bdeb(?:-src)?\b", block):
                continue
            block = re.sub(r"(?m)^Types:.*$", "Types: deb-src", block)
            blocks.append(block)
            for snapshot in re.findall(r"(?m)^#\s*(https?://snapshot\.debian\.org/archive/[^\s]+)", block):
                historic = re.sub(r"(?m)^URIs:.*$", "URIs: " + snapshot, block)
                historic = re.sub(r"(?m)^Check-Valid-Until:.*\n?", "", historic)
                blocks.append(historic + "\nCheck-Valid-Until: no")
        converted = "\n\n".join(dict.fromkeys(blocks)) + "\n"
        uri_fields = re.findall(r"(?m)^URIs:([^\n]*(?:\n[ \t]+[^\n]*)*)", converted)
        uris = [uri for field in uri_fields for uri in field.split()]
    require(not converted.strip() or uris, "APT source configuration has no archive URI")
    for uri in uris:
        require_debian_uri(uri)
    return converted


def uri_records(text):
    records = {}
    for line in text.splitlines():
        if not line.startswith("'"):
            continue
        fields = shlex.split(line)
        require(len(fields) == 4 and fields[2].isdecimal(), "Malformed apt source URI output")
        uri, name, size, _ = fields
        require_debian_uri(uri)
        require(Path(name).name == name and name not in records, "Unsafe or repeated apt source file")
        records[name] = {"url": uri, "size": int(size)}
    require(records, "APT produced no source download URIs")
    return records


def verify_package(directory, name, version, expected):
    for filename, info in expected.items():
        path = directory / filename
        require(path.is_file() and path.stat().st_size == info["size"] and digest(path) == info["sha256"],
                f"Incomplete or mismatched source payload: {name}={version}/{filename}")
    dsc = next(directory / n for n in expected if n.endswith(".dsc"))
    records = controls(dsc.read_text())
    require(len(records) == 1 and records[0].get("Source") == name and records[0].get("Version") == version,
            f"Wrong source tuple in {dsc.name}")
    payloads = checksums(records[0].get("Checksums-Sha256", ""))
    require(payloads == {n: v for n, v in expected.items() if not n.endswith(".dsc")},
            f".dsc and archive index disagree for {name}={version}")


def manifest_digest(image_info):
    descriptor = image_info.get("Descriptor", {})
    require(isinstance(descriptor, dict) and descriptor.get("mediaType") in MANIFEST_MEDIA_TYPES,
            "Docker must expose a platform image manifest descriptor, not an index or legacy image ID")
    value = descriptor.get("digest", "")
    require(isinstance(value, str) and value.startswith("sha256:") and SHA256.fullmatch(value[7:]),
            "Invalid platform manifest digest")
    return value


def reuse_source_package(cache, destination, name, version, expected):
    """Read a prior tuple only after authenticating its exact files against the fresh index."""
    if cache is None:
        return False
    source = cache / name / version
    if not source.exists() and not source.is_symlink():
        return False
    require(source.is_dir() and not source.is_symlink(), "Existing source-cache tuple is not a regular directory")
    verify_package(source, name, version, expected)
    for filename in expected:
        shutil.copyfile(source / filename, destination / filename)
    verify_package(destination, name, version, expected)
    return True


def prepare(args):
    output = args.output.resolve()
    require(not output.exists() or not any(output.iterdir()), "Output must be absent or empty; failed evidence is not overwritten")
    output.mkdir(parents=True, exist_ok=True)
    content = output / "content"
    content.mkdir()
    log = (output / "collector.log").open("w")
    containers = []
    deadline = time.monotonic() + 1800

    def run(command, timeout=180):
        remaining = deadline - time.monotonic()
        require(remaining > 0, "Source collection exceeded its 30-minute deadline")
        log.write(shlex.join(command) + "\n")
        log.flush()
        process = subprocess.run(command, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                 timeout=min(timeout, remaining))
        require(len(process.stdout) <= 20 * 1024**2 and len(process.stderr) <= 1024**2,
                "Unexpectedly large collector command output")
        log.write(process.stdout + process.stderr)
        log.flush()
        require(process.returncode == 0, f"Command failed ({process.returncode}); inspect {output / 'collector.log'}")
        return process.stdout

    try:
        manifest = {"version": args.version, "vendorSourceCommit": args.vendor_source_commit,
                    "vendorSourceSha256": args.vendor_source_sha256, "images": [], "sources": []}
        repository = Path(__file__).resolve().parents[1]
        manifest["gridgrindBuildSources"] = copy_build_inputs(repository, content)
        tuples = set()
        configs = {}
        for platform, image in args.image:
            name = "gridgrind-source-" + uuid.uuid4().hex
            run(["docker", "create", "--pull=never", "--platform", platform, "--user", "0",
                 "--name", name, "--mount", f"type=bind,source={output},target=/source-output",
                 "--entrypoint", "sh", image, "-c", "while :; do sleep 3600; done"])
            containers.append(name)
            run(["docker", "start", name])
            actual_arch = run(["docker", "exec", name, "dpkg", "--print-architecture"]).strip()
            require(actual_arch == PLATFORMS[platform], f"Wrong architecture for {platform}: {actual_arch}")
            package_text = run(["docker", "exec", name, "dpkg-query", "-W", "-f",
                                "${db:Status-Status}\t${binary:Package}\t${Version}\t${source:Package}\t${source:Version}\n"])
            image_dir = content / "images" / actual_arch
            image_dir.mkdir(parents=True)
            (image_dir / "dpkg-status.tsv").write_text(package_text)
            packages = []
            for line in package_text.splitlines():
                fields = line.split("\t")
                require(len(fields) == 5, "Incomplete dpkg source mapping")
                if fields[0] != "installed":
                    continue
                binary, binary_version, source, source_version = fields[1:]
                require(all(fields[1:]), "Incomplete installed package source mapping")
                require(SOURCE_NAME.fullmatch(source), "Invalid Debian source package name")
                require(SOURCE_VERSION.fullmatch(source_version), "Invalid Debian source version")
                tuples.add((source, source_version))
                packages.append({"binary": binary, "version": binary_version, "source": source, "sourceVersion": source_version})
            (image_dir / "packages.tsv").write_text("".join(
                "\t".join([p["binary"], p["version"], p["source"], p["sourceVersion"]]) + "\n" for p in packages))
            release = run(["docker", "exec", name, "sh", "-ec", 'cat "$JAVA_HOME/release"'])
            commit = run(["docker", "exec", name, "sh", "-ec", 'cat "$JAVA_HOME/commitId.txt"']).strip()
            require(commit == args.vendor_source_commit, f"Vendor source commit does not match {platform} JRE")
            (image_dir / "java-release.txt").write_text(release)
            (image_dir / "java-commitId.txt").write_text(commit + "\n")
            image_id = json.loads(run(["docker", "inspect", name]))[0]["Image"]
            image_info = json.loads(run(["docker", "image", "inspect", "--platform", platform, image_id]))[0]
            manifest["images"].append({"platform": platform, "manifestDigest": manifest_digest(image_info), "packages": packages})
            paths = run(["docker", "exec", name, "sh", "-ec",
                         "find /etc/apt -maxdepth 2 -type f \\( -name '*.sources' -o -name '*.list' \\) -print"])
            for source_path in paths.splitlines():
                text = run(["docker", "exec", name, "cat", source_path])
                filename = Path(source_path).name
                (image_dir / filename).write_text(text)
                converted = source_config(text, filename)
                if converted.strip():
                    configs.setdefault((converted, Path(filename).suffix), None)

        require(tuples and configs, "Final images have no Debian source mapping or apt archive configuration")
        helper = containers[0]
        config_dir = content / "apt-source-config"
        config_dir.mkdir()
        for index, ((text, suffix), _) in enumerate(configs.items()):
            (config_dir / f"archive-{index}{suffix}").write_text(text)
        apt_options = ["-o", "Dir::Etc::sourcelist=/dev/null", "-o", "Dir::Etc::sourceparts=/source-output/content/apt-source-config",
                       "-o", "Dir::State::lists=/tmp/gridgrind-source-lists", "-o", "Dir::Cache=/tmp/gridgrind-source-cache",
                       "-o", "Acquire::Retries=2", "-o", "Acquire::http::Timeout=60", "-o", "Acquire::https::Timeout=60"]
        run(["docker", "exec", helper, "mkdir", "-p", "/tmp/gridgrind-source-lists/partial", "/tmp/gridgrind-source-cache/archives/partial"])
        run(["docker", "exec", helper, "apt-get", *apt_options, "update"], timeout=600)
        index_size = int(run(["docker", "exec", helper, "du", "-sb", "/tmp/gridgrind-source-lists"]).split()[0])
        require(index_size <= INDEX_LIMIT, "Debian source index evidence exceeds 256 MiB")
        plans = []
        for name, version in sorted(tuples):
            text = run(["docker", "exec", helper, "apt-cache", *apt_options, "showsrc", name])
            expected = source_record(text, name, version)
            uris = uri_records(run(["docker", "exec", helper, "apt-get", *apt_options,
                                    "--print-uris", "--download-only", "--only-source", "source", f"{name}={version}"]))
            require(uris.keys() == expected.keys() and all(uris[n]["size"] == expected[n]["size"] for n in expected),
                    f"APT source URI set disagrees with authenticated metadata for {name}={version}")
            plans.append({"package": name, "version": version, "files": expected, "uris": uris, "indexStanza": text})
        planned = (sum(v["size"] for p in plans for v in p["files"].values())
                   + args.vendor_source.stat().st_size + index_size
                   + sum(source["size"] for source in manifest["gridgrindBuildSources"]))
        # Conservative publication peak: upstream tree, release archive, BuildKit upload
        # snapshot, source-image layer, Docker download blob, and unpacked layer.
        # Existing runtime-image caches have already consumed their earlier disk budget.
        required_space = 6 * planned + 512 * 1024**2
        free_space = shutil.disk_usage(output).free
        print(f"Preflight: {len(plans)} Debian source tuples; {planned} planned source/index bytes; "
              f"required free bytes={required_space}; available free bytes={free_space} "
              "(six preparation/publication copies plus 512 MiB headroom)", flush=True)
        require(planned + 16 * 1024**2 < ASSET_LIMIT, "Source payload would exceed the single GitHub release asset limit")
        require(free_space >= required_space, "Insufficient free space for source preparation and registry publication")
        (output / "source-plan.json").write_text(json.dumps(plans, indent=2) + "\n")
        for plan in plans:
            name, version = plan["package"], plan["version"]
            directory = content / "debian" / name / version
            directory.mkdir(parents=True)
            relative = directory.relative_to(output)
            if reuse_source_package(args.source_cache, directory, name, version, plan["files"]):
                print(f"Reused verified source tuple {name}={version}", flush=True)
            else:
                run(["docker", "exec", "--workdir", "/source-output/" + str(relative), helper,
                     "apt-get", *apt_options, "--download-only", "--only-source", "source", f"{name}={version}"], timeout=900)
            verify_package(directory, name, version, plan["files"])
            (directory / "source-index.txt").write_text(plan["indexStanza"])
            manifest["sources"].append({k: v for k, v in plan.items() if k != "indexStanza"})
        trust = content / "apt-index-evidence"
        trust.mkdir()
        paths = run(["docker", "exec", helper, "find", "/tmp/gridgrind-source-lists", "-maxdepth", "1", "-type", "f"])
        for path in paths.splitlines():
            if any(token in Path(path).name for token in ("InRelease", "Release", "Sources")):
                run(["docker", "cp", helper + ":" + path, str(trust / Path(path).name)])
        vendor = content / "vendor"
        vendor.mkdir()
        vendor_archive = vendor / f"corretto-{args.vendor_source_commit}.tar.gz"
        shutil.copyfile(args.vendor_source, vendor_archive)
        require(digest(vendor_archive) == args.vendor_source_sha256, "Copied vendor source hash changed")
        manifest["vendorSourceArchive"] = vendor_archive.relative_to(content).as_posix()
        for source in manifest["gridgrindBuildSources"]:
            require(digest(repository / source["sourcePath"]) == source["sha256"],
                    "Container build source changed during collection: " + source["sourcePath"])
        (content / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
        (content / "README.txt").write_text(
            "Corresponding source for the listed GridGrind container platform manifests.\n"
            "manifest.json maps both image platforms and Debian binary packages to exact source versions.\n"
            "Each debian/<package>/<version> directory retains the original .dsc and every source part.\n"
            "Extract with: dpkg-source -x <package>.dsc (Debian dpkg-dev). Debian packaging/build rules\n"
            "and component license notices are inside the source archives. No rebuilding is claimed here.\n"
            "vendor/ contains the unchanged complete vendor commit archive identified in the manifest;\n"
            "its source, native code, build scripts and licenses retain their upstream terms.\n"
            "gridgrind-build/ contains the exact image Dockerfile, runtime pins and source collector,\n"
            "with their hashes in manifest.json. Those GridGrind files use MPL2; their LICENSE and\n"
            "NOTICE preserve the project scope, attribution and source-availability information.\n"
            "The corresponding GridGrind release tag separately publishes the application source.\n"
            "This is delivered source material, not a new written offer for later source delivery.\n")
        archive = output / f"gridgrind-container-sources-{args.version}.tar.gz"
        write_archive(content, archive, args.version)
        require(archive.stat().st_size <= ASSET_LIMIT, "Prepared source archive exceeds GitHub release asset limit")
        public_manifest = output / f"gridgrind-container-sources-{args.version}.manifest.json"
        shutil.copyfile(content / "manifest.json", public_manifest)
        (output / f"gridgrind-container-sources-{args.version}.sha256").write_text(
            "".join(digest(path) + "  " + path.name + "\n" for path in [archive, public_manifest]))
        print(archive, flush=True)
    finally:
        primary_failure = sys.exc_info()[0] is not None
        cleanup_failures = []
        for container in containers:
            # Source files belong to the invoking host user even when the helper runs as root.
            for command in (["docker", "exec", container, "chown", "-R", f"{os.getuid()}:{os.getgid()}", "/source-output"],
                            ["docker", "rm", "-f", container]):
                try:
                    result = subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60)
                    if result.returncode:
                        cleanup_failures.append(shlex.join(command))
                except (OSError, subprocess.TimeoutExpired) as error:
                    cleanup_failures.append(str(error))
        for failure in cleanup_failures:
            print("Source collector cleanup failed: " + failure, file=sys.stderr)
        log.close()
        require(primary_failure or not cleanup_failures, "Could not clean source collector helpers; inspect stderr")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--image", action="append", required=True, metavar="PLATFORM=LOCAL_IMAGE")
    parser.add_argument("--version", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--source-cache", type=Path,
                        help="Read-only prior content/debian directory; reuse requires fresh authenticated hashes")
    parser.add_argument("--vendor-source", type=Path, required=True)
    parser.add_argument("--vendor-source-sha256", required=True)
    parser.add_argument("--vendor-source-commit", required=True)
    args = parser.parse_args()
    try:
        require(re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", args.version), "Version must be X.Y.Z")
        args.image = [tuple(item.split("=", 1)) for item in args.image]
        require(all(len(item) == 2 and item[0] in PLATFORMS and item[1] for item in args.image), "Invalid platform=image input")
        require(sorted(p for p, _ in args.image) == sorted(PLATFORMS), "Supply exactly amd64 and arm64 final images")
        require(SHA256.fullmatch(args.vendor_source_sha256), "Invalid vendor source SHA256")
        require(re.fullmatch(r"[a-f0-9]{40}", args.vendor_source_commit), "Vendor source commit must be a full SHA1")
        require(args.vendor_source.is_file() and digest(args.vendor_source) == args.vendor_source_sha256,
                "Vendor source file is missing or its SHA256 differs")
        if args.source_cache is not None:
            args.source_cache = args.source_cache.resolve()
            require(args.source_cache.is_dir(), "Source cache must be an existing prior content/debian directory")
        def interrupt(*_):
            raise KeyboardInterrupt()

        signal.signal(signal.SIGTERM, interrupt)
        prepare(args)
    except (ValueError, OSError, subprocess.TimeoutExpired) as error:
        parser.exit(1, f"Source preparation failed: {error}\n")
    except KeyboardInterrupt:
        parser.exit(130, "Source preparation interrupted; partial evidence retained.\n")


if __name__ == "__main__":
    main()
