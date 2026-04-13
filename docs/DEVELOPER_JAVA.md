---
afad: "3.5"
version: "0.41.0"
domain: DEVELOPER_JAVA
updated: "2026-04-11"
route:
  keywords: [gridgrind, java26, jdk, jdk.java.net, shell, zsh, java-home, javac, brew, gradlew, macos]
  questions: ["how is java 26 set up for gridgrind", "why does gridgrind require shell-level java 26", "how do i configure local java 26 for gridgrind", "why is /usr/bin/java wrong for gridgrind", "why should i use ./gradlew instead of brew gradle"]
---

# Java 26 Developer Setup

**Purpose**: Document the local Java 26 setup used for GridGrind development, why it is configured this way, and how to maintain it safely.
**Prerequisites**: macOS with zsh.

## Current Stance

GridGrind targets Java 26.

That target is enforced in project configuration, GitHub Actions, Docker, and the packaged CLI, so
local developer shells should also resolve to Java 26 by default.

The current local setup intentionally uses:
- the official Java 26 release from [jdk.java.net/26](https://jdk.java.net/26/)
- a user-local installation under `~/Library/Java/JavaVirtualMachines/openjdk-26.jdk`
- explicit `JAVA_HOME` and `PATH` wiring in zsh startup files
- no Homebrew-managed JDK on the machine
- the Gradle wrapper (`./gradlew`) as the only supported Gradle entrypoint

## Why This Setup

Reasons:
- `java -jar gridgrind.jar` uses the ambient shell `java`, not Gradle toolchains.
- `./check.sh` and release work need the launcher runtime to match the project baseline.
- macOS still ships the `/usr/bin/java` launcher stub, which is the wrong runtime for GridGrind.
- GitHub Actions and the Docker release surface already target Java 26 explicitly, so local shells
  should match that contract.
- `jdk.java.net` gives a clear, upstream, version-specific source instead of waiting on a package
  manager mirror cadence.
- the Gradle wrapper pins the repository's actual Gradle version, while Brew `gradle` introduces an
  unnecessary second moving part and can pull `openjdk` back in as a dependency.

## Installed Layout

Current local bundle:
- `~/Library/Java/JavaVirtualMachines/openjdk-26.jdk`

Current upstream artifact used for this setup:
- archive: `openjdk-26_macos-aarch64_bin.tar.gz`
- source page: [jdk.java.net/26](https://jdk.java.net/26/)
- published SHA-256 at time of install:
  `254586bcd1bf6dcd125ad667ac32562cb1e2ab1abf3a61fb117b6fabb571e765`

Current shell resolution:
- `JAVA_HOME=/Users/erst/Library/Java/JavaVirtualMachines/openjdk-26.jdk/Contents/Home`
- `java` resolves to `${JAVA_HOME}/bin/java`
- `javac` resolves to `${JAVA_HOME}/bin/javac`

The old Homebrew-linked symlink was intentionally removed:
- `~/Library/Java/JavaVirtualMachines/openjdk.jdk`

## Shell Configuration

The current shell setup is intentionally duplicated across login and interactive zsh startup.

In [~/.zprofile](/Users/erst/.zprofile):

```zsh
eval "$(/opt/homebrew/bin/brew shellenv zsh)"

if /usr/libexec/java_home -v 26 >/dev/null 2>&1; then
    export JAVA_HOME="$("/usr/libexec/java_home" -v 26)"
    case ":${PATH}:" in
        *":${JAVA_HOME}/bin:"*) ;;
        *) export PATH="${JAVA_HOME}/bin:${PATH}" ;;
    esac
fi
```

In [~/.zshrc](/Users/erst/.zshrc):

```zsh
if [[ -z "${JAVA_HOME:-}" ]] && /usr/libexec/java_home -v 26 >/dev/null 2>&1; then
  export JAVA_HOME="$("/usr/libexec/java_home" -v 26)"
fi
if [[ -n "${JAVA_HOME:-}" ]]; then
  case ":${PATH}:" in
    *":${JAVA_HOME}/bin:"*) ;;
    *) export PATH="${JAVA_HOME}/bin:${PATH}" ;;
  esac
fi
```

Why both files:
- `~/.zprofile` covers login shells
- `~/.zshrc` covers interactive non-login shells
- this avoids one terminal path seeing Java 26 while another silently falls back to a different JDK

## Installation Procedure

This is the exact local procedure used for GridGrind.

1. Download the official macOS AArch64 Java 26 archive from [jdk.java.net/26](https://jdk.java.net/26/).
2. Verify the archive checksum against the value published by the OpenJDK site.
3. Extract the bundle into `~/Library/Java/JavaVirtualMachines/openjdk-26.jdk`.
4. Add the zsh configuration shown above so `JAVA_HOME` and `PATH` prefer Java 26.
5. Remove stale JDK symlinks that still point to Homebrew-managed Java, especially `~/Library/Java/JavaVirtualMachines/openjdk.jdk`.
6. Verify both shell modes:
   `zsh -ic 'echo $JAVA_HOME; command -v java; command -v javac'`
   and
   `zsh -lic 'echo $JAVA_HOME; command -v java; command -v javac'`

## Build Logic Caveat

GridGrind's product modules and runtime baseline are Java 26.

There is one deliberate exception: the shared included build logic under `gradle/build-logic`
still emits JVM 25 bytecode because Kotlin `2.3.0` does not yet target JVM 26 directly. That
build now compiles with the Java 26 toolchain and only lowers the emitted bytecode level, so no
separate Java 25 installation is required. That does **not** change the project baseline:
- the Gradle launcher JVM is Java 26
- the product modules target Java 26
- the CLI fat JAR requires Java 26
- `./check.sh` now fails fast if the active shell does not already resolve to Java 26

Treat shared build-logic JVM 25 as a temporary toolchain boundary, not as permission to run
GridGrind on a Java 25 shell.

## Homebrew Cleanup

The old Homebrew JDK was intentionally removed.

Important detail:
- Homebrew can refuse the uninstall because the `gradle` formula depends on `openjdk`
- if that happens and the repo is already green through `./gradlew`, the acceptable cleanup is:
  `brew uninstall --ignore-dependencies openjdk`

Why that is acceptable:
- GridGrind uses the project Gradle wrapper as the canonical entrypoint
- the shell resolves Java 26 directly through `JAVA_HOME`
- the Homebrew JDK is redundant and is the main source of local version ambiguity

## Pitfalls

Known pitfalls:
- macOS `/usr/bin/java` is a launcher stub, not the project runtime. If `command -v java` returns
  `/usr/bin/java`, stop and fix the shell before running GridGrind.
- `java -jar` does not use Gradle toolchains. That is why shell-level Java 26 matters even though
  Gradle can provision toolchains for compilation.
- `~/.zprofile` alone is not enough. Some terminal launches use interactive non-login shells and
  would otherwise miss the Java 26 override.
- stale symlinks under `~/Library/Java/JavaVirtualMachines/` can silently point back to removed or
  outdated JDKs.
- if `brew shellenv` is evaluated after Java path setup, Homebrew paths can regain precedence.
- some automation environments start non-interactive shells and do not load zsh startup files. In
  those cases, inspect `command -v java` before trusting the runtime.
- Brew `gradle` can reintroduce `openjdk` dependency churn. Use `./gradlew` instead.

## Verification Commands

Useful checks:

```bash
/usr/libexec/java_home -V
command -v java
command -v javac
java --version
zsh -ic 'echo $JAVA_HOME; command -v java; command -v javac'
zsh -lic 'echo $JAVA_HOME; command -v java; command -v javac'
sed -n '1,20p' "$(/usr/libexec/java_home -v 26)/release"
```

Expected outcomes:
- Java 26 appears in `/usr/libexec/java_home -V`
- `command -v java` does not point at `/usr/bin/java`
- both zsh modes resolve `java` and `javac` inside `openjdk-26.jdk`
- the `release` file reports `JAVA_VERSION="26"`

## Full Toolchain Verification

After Java 26 is active in the shell, the canonical GridGrind verification sequence is:

```bash
java --version
./gradlew --version --console=plain
./gradlew check --no-daemon --console=plain
./check.sh --console=plain
./scripts/docker-smoke.sh
```

Run those commands from a real terminal session.

Why:
- local shell startup files are what make Java 26 the default `java`
- non-interactive wrappers may skip `~/.zprofile` and `~/.zshrc`
- GridGrind release work, `java -jar`, and the CLI all depend on the ambient runtime, not just on
  Gradle toolchains

## Maintenance Guidance

When Java 27 or later becomes the project target:
- install the new upstream JDK beside the old one first
- update the zsh config to use the new major version
- verify both shell modes before removing the previous JDK
- update GridGrind's shell docs, build scripts, CI, Docker, and release surfaces together
