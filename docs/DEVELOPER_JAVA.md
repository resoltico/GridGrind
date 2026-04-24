---
afad: "3.5"
version: "0.59.0"
domain: DEVELOPER_JAVA
updated: "2026-04-25"
route:
  keywords: [gridgrind, java26, gradle wrapper, java_home, macos, shell, mounted volumes, local disk]
  questions: ["what java setup does gridgrind actually require", "how should i configure java 26 for gridgrind", "does the jdk vendor matter for gridgrind", "why does gridgrind need java in the shell"]
---

# Java 26 And Gradle Workstation Setup

**Purpose**: Document the Java and Gradle setup that GridGrind actually enforces for contributors.
**Prerequisites**: macOS with zsh.
**Companion references**: [DEVELOPER.md](./DEVELOPER.md), [DEVELOPER_GRADLE.md](./DEVELOPER_GRADLE.md)

## What The Repository Actually Requires

GridGrind's product modules target Java 26 and the repository build is wrapper-first.

What the codebase and scripts actually enforce today:

- the active shell must resolve a full JDK 26, not only a JRE
- `java` and `javac` must both be reachable from the shell
- `java --version` and `javac --version` must both report major version `26`
- if `java` or `javac` resolve through macOS `/usr/bin` launchers, those launchers must
  immediately resolve to the intended installed JDK instead of triggering the Apple install stub
- all repository build and test commands should use `./gradlew`
- full verification should run from a local-disk checkout or local mirror, not from a mounted
  external volume on macOS

The repository does **not** enforce one specific JDK vendor. CI currently uses Azul Zulu 26
through `actions/setup-java`, but any compatible JDK 26 that satisfies the checks above is valid.

## Canonical Stance

| Component | Recommended source | Why |
|:----------|:-------------------|:----|
| JDK | any compatible JDK 26 distribution | GridGrind validates the major version and shell visibility, not the vendor string |
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

These flows still depend on the shell-visible `java` command:

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

## Verification Commands

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

## Full Repository Verification

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

- relying on the Apple install-stub behavior behind `/usr/bin/java` or `/usr/bin/javac` instead
  of a real JDK 26 launcher
- assuming Gradle toolchains cover `java -jar` or `./check.sh`
- using `gradle` from `PATH` for repo work instead of `./gradlew`
- running full verification from a mounted external volume on macOS
- assuming the JDK vendor must match local docs exactly when the real invariant is Java 26 plus a
  shell-visible full JDK
