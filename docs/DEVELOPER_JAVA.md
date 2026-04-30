---
afad: "4.0"
version: "0.62.0"
domain: DEVELOPER_JAVA
updated: "2026-05-01"
route:
  keywords: [gridgrind, java26, zulu26, gradle wrapper, java_home, macos, shell, devcontainer, mounted volumes, local disk]
  questions: ["what java setup does gridgrind actually require", "how should i configure java 26 for gridgrind", "do i need java installed on the host if i use the devcontainer", "does the jdk vendor matter for gridgrind", "why does gridgrind need java in the shell"]
---

# Java 26 And Gradle Workstation Setup

**Purpose**: Document the Java setup GridGrind enforces for host-native contributors and the Java
alignment the preferred devcontainer path already bakes in.
**Prerequisites**: macOS with zsh.
**Companion references**: [DEVELOPER.md](./DEVELOPER.md),
[DEVELOPER_DEVCONTAINER.md](./DEVELOPER_DEVCONTAINER.md), [DEVELOPER_GRADLE.md](./DEVELOPER_GRADLE.md)

## Preferred Path First

If you use the committed devcontainer, you do not need Java 26 installed in the host shell for
normal repository work. The container already ships Azul Zulu 26 and is the preferred contributor
path for this repository, whether you enter it through VS Code or through the Dev Container CLI.

This file matters in two cases:

- you are deliberately using the host-native contributor path
- you want to understand the Java contract the devcontainer and CI are aligned to

## What The Repository Actually Requires

GridGrind's product modules target Java 26 and the repository build is wrapper-first.

What the codebase and scripts actually enforce today for host-native work:

- the active shell must resolve a full JDK 26, not only a JRE
- `java` and `javac` must both be reachable from the shell
- `java --version` and `javac --version` must both report major version `26`
- if `java` or `javac` resolve through macOS `/usr/bin` launchers, those launchers must
  immediately resolve to the intended installed JDK instead of triggering the Apple install stub
- all repository build and test commands should use `./gradlew`
- full verification should run from a local-disk checkout or local mirror, not from a mounted
  external volume on macOS

The repository runtime contract does **not** require one specific host JDK vendor, but the project
now standardizes CI and the preferred devcontainer on Azul Zulu 26. Host-native work may use any
compatible JDK 26 that satisfies the checks above, though Zulu 26 is the preferred host-native
match when you want the closest contributor-to-CI alignment.

## Canonical Stance

| Component | Recommended source | Why |
|:----------|:-------------------|:----|
| Preferred contributor Java | devcontainer-baked Azul Zulu 26 | matches CI vendor and keeps Java tooling out of the host process tree |
| Host-native JDK | Azul Zulu 26 when convenient; any compatible JDK 26 otherwise | GridGrind validates the major version and shell visibility, while Zulu keeps host-native behavior closest to CI and the devcontainer |
| Gradle | repository `./gradlew` | the wrapper pins the project Gradle version and is the only supported repo entrypoint |
| Global `gradle` | optional but unused for repo work | GridGrind already commits wrapper files |

Official Java references that are still useful when picking or verifying a JDK 26 install:

- [OpenJDK JDK 26 project page](https://openjdk.org/projects/jdk/26/)
- [OpenJDK install guidance](https://openjdk.org/install/index.html)
- [Archived OpenJDK GA builds](https://jdk.java.net/archive/)

Do not hardcode one specific archive filename or checksum in local docs: those details move as the
published binaries and archives move forward.

## Why Shell Java Still Matters

Gradle toolchains help the build compile consistently, but they do not replace the ambient shell
runtime for every contributor path.

These flows still depend on the shell-visible `java` command when you are outside the devcontainer:

- `java -jar cli/build/libs/gridgrind.jar`
- `./check.sh`
- local release-surface verification and Docker smoke setup
- ad hoc CLI checks such as `./gradlew :cli:run --args="--version"`

That is why the repository docs focus on shell-visible Java 26 rather than only on Gradle
configuration.

## Local Filesystem Rule

GridGrind's full verification expects a local filesystem with working file-lock semantics.

Why:

- Gradle project cache and JaCoCo execution data both rely on file locking
- mounted external volumes on macOS can fail with errors such as `Operation not supported`
- `./gradlew check` or even earlier build stages can fail before product code is exercised if the
  workspace lives on the wrong filesystem

If your main checkout lives on a mounted external volume, keep that checkout for editing if you
want, but do the full build/test loop from a local-disk worktree or a disposable local mirror.
The canonical `./check.sh` gate also runs with `--no-daemon`, uses a repo-scoped
`GRADLE_USER_HOME` under `tmp/gradle-user-home` by default, and shares one repo-wide verification
lock with the Docker and Jazzer top-level entrypoints.

## Typical macOS Shell Setup

If you use `/usr/libexec/java_home`, a minimal zsh setup looks like this.

In `~/.zprofile`:

```zsh
if /usr/libexec/java_home -v 26 >/dev/null 2>&1; then
  export JAVA_HOME="$('/usr/libexec/java_home' -v 26)"
  path=("${JAVA_HOME}/bin" ${path:#${JAVA_HOME}/bin})
fi
```

In `~/.zshrc`:

```zsh
if [[ -z "${JAVA_HOME:-}" ]] && /usr/libexec/java_home -v 26 >/dev/null 2>&1; then
  export JAVA_HOME="$('/usr/libexec/java_home' -v 26)"
fi
if [[ -n "${JAVA_HOME:-}" ]]; then
  path=("${JAVA_HOME}/bin" ${path:#${JAVA_HOME}/bin})
fi
```

The goal is simple: both login and interactive shells should resolve the same JDK 26 first.

## Host-Native Verification Commands

Run these from the same shell you will use for repo work:

```bash
echo "$JAVA_HOME"
command -v java
command -v javac
java --version
javac --version
./gradlew --version --console=plain
```

Expected outcomes:

- `java --version` and `javac --version` report major version `26`
- if `command -v java` or `command -v javac` point at `/usr/bin/java` or `/usr/bin/javac` on
  macOS, those launchers still resolve immediately to the intended installed JDK instead of
  surfacing the Apple install stub
- `./gradlew --version` reports the wrapper-pinned Gradle version

## Host-Native Full Repository Verification

Once Java 26 is active in the shell, the normal GridGrind verification loop is:

```bash
java --version
./gradlew --version --console=plain
./gradlew check --console=plain
./check.sh --console=plain
```

If the repository is on a mounted volume and the build fails early with filesystem-locking errors,
rerun that sequence from a local-disk mirror or worktree instead of weakening the normal build.

## Common Pitfalls

- forgetting that the preferred devcontainer path already provides Java 26 and then spending time
  debugging a host Java install you do not actually need
- relying on the Apple install-stub behavior behind `/usr/bin/java` or `/usr/bin/javac` instead
  of a real JDK 26 launcher
- assuming Gradle toolchains cover `java -jar` or `./check.sh`
- using `gradle` from `PATH` for repo work instead of `./gradlew`
- running full verification from a mounted external volume on macOS
- assuming the JDK vendor must match local docs exactly when the real invariant is Java 26 plus a
  shell-visible full JDK
