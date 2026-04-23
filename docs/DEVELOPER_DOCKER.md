---
afad: "3.5"
version: "0.55.0"
domain: DEVELOPER_DOCKER
updated: "2026-04-16"
route:
  keywords: [gridgrind, docker, docker desktop, docker smoke, check.sh, anonymous docker config, docker context, container]
  questions: ["how should i set up docker for gridgrind", "why should gridgrind use an anonymous docker config for docker smoke", "what docker runtime is supported for gridgrind", "how do i verify docker before running check.sh"]
---

# Docker Workstation Setup

**Purpose**: Codify the supported Docker setup for GridGrind contributors on macOS.
**Prerequisites**: Java 26 and wrapper-based Gradle setup already in place through
[DEVELOPER_JAVA.md](./DEVELOPER_JAVA.md).

Supported workstation shape:
- Docker Desktop installed from Docker's own macOS distribution path on `docker.com`
- the Docker daemon is running and reachable from the current shell
- the `docker buildx` plugin is available in the current shell
- the active Docker context targets the local Docker Desktop engine
- the repository checkout lives on the Mac's local filesystem

## Canonical Stance

For GridGrind's local container work, the documented standard is:
- Docker comes from Docker Desktop, not from a separate Homebrew-only container-runtime story
- `docker` and the Docker daemon must already work in the current shell before `./check.sh`
- `docker buildx` is required; the smoke gate uses `docker buildx build --load`, not Docker's
  deprecated legacy builder path
- local smoke and release verification must not depend on personal Docker login state
- public-image verification should run through a temporary anonymous `DOCKER_CONFIG` while still
  targeting the active local Docker engine
- any mounted-path fixture files created by smoke scripts must keep the same ownership and
  filesystem semantics as real operator output instead of introducing weaker test-only behavior

The repository now enforces these Docker-runtime rules in `scripts/docker-smoke.sh` and
`scripts/verify-container-publication.sh`.

Release-build reproducibility is part of that contract too:
- `Dockerfile` pins the Azul Java 26 base image to a manifest-list digest instead of a floating
  tag; update that digest deliberately when the runtime base moves forward
- the production image must keep the minimal headless font stack required by signature-line
  preview generation (`fontconfig` plus DejaVu today) so Docker matches the fat-JAR drawing
  surface instead of silently dropping `SET_SIGNATURE_LINE`
- the GHCR publication workflow emits OCI provenance and SBOM attestations for the published
  multi-arch image in addition to the runnable image tags

## Why Anonymous Docker Config Matters

GridGrind's Docker smoke and release verification depend on public base-image or published-image
fetches. Those operations should not depend on:
- Docker Desktop credential-helper availability
- a contributor's personal Docker Hub login state
- Docker Desktop plugin and hook behavior in `~/.docker/config.json`

On fresh macOS machines, Docker Desktop's credential helper can stall public metadata fetches even
though the daemon itself is healthy. GridGrind's Docker verification therefore uses a temporary
empty `DOCKER_CONFIG`, derives the active engine endpoint from the current Docker context, and only
if that empty config would hide Buildx, stages an already-installed host `docker-buildx` plugin
into the anonymous config. On macOS that plugin often comes from Docker Desktop; on CI or other
hosts it may come from a system CLI-plugin directory. That keeps public pulls, runs, and Buildx
image loads independent from personal Docker auth state without falling back to Docker's
deprecated legacy builder path.

## Verification

Before running GridGrind's whole-repo gate, confirm the shell sees a live Docker runtime:

```bash
docker --version
docker buildx version
docker context show
docker info --format '{{.ServerVersion}}'
```

Expected local shape on Docker Desktop:
- `docker --version` returns a real Docker CLI version
- `docker buildx version` returns a real Buildx version
- `docker info` returns a server version instead of a connection error
- `docker context show` usually prints `desktop-linux`

Then the supported local gates are:

```bash
./scripts/docker-smoke.sh
./check.sh
```

`./check.sh` Stage 5 invokes `scripts/docker-smoke.sh`, which:
- builds the local image from the repository root through `docker buildx build --load`
- runs mounted-path container commands under the caller's UID:GID so response files and saved
  workbooks stay owned by the invoking operator on both macOS Docker Desktop and Linux CI runners
- verifies `--help` and `--version`
- verifies create-from-`NEW`, reopen-from-`EXISTING`, signature-line authoring, and
  `STREAMING_WRITE` readback flows through a mounted working directory with spaces and punctuation
  in the paths
- rejects unexpected stderr on the successful request paths so agent consumers do not silently
  accumulate logger or runtime-noise regressions

## Troubleshooting

If Docker verification fails on a fresh machine:
- confirm Docker Desktop is actually running, not only installed
- rerun `docker buildx version`, `docker info`, and `docker context show` from the same shell that
  will run `./check.sh`
- prefer fixing the local Docker runtime over weakening `./check.sh`
- if a public pull still hangs, inspect whether personal Docker config customizations were
  reintroduced into the verification path
- if mounted-path response or workbook files come back with unexpected ownership, inspect whether
  the container invocation is still running under the caller UID:GID before weakening the smoke
  contract
