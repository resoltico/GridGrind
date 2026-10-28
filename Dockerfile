# syntax=docker/dockerfile:1.28.0@sha256:bb22d9815c728170f72750f4e5b0d672e06176142e1d602c7e66c050100b7e5b

# Pin both builder and runtime manifest lists so local rebuilds and published images stay
# reproducible across time.
FROM --platform=$BUILDPLATFORM azul-zulu:27.30-27.0.1-jdk-debian13@sha256:7f3f17771a7a8500b8844c7da21b89f14feb856b7bafe6d68292bca8e746f035 AS build

ARG GRIDGRIND_VERSION

WORKDIR /workspace

COPY gradlew gradle.properties settings.gradle.kts build.gradle.kts gradle.lockfile settings-gradle.lockfile ./
COPY gradle ./gradle
COPY LICENSE NOTICE PATENTS.md LICENSE-APACHE-2.0 LICENSE-BSD-2-CLAUSE LICENSE-BSD-3-CLAUSE ./
COPY cli ./cli
COPY core ./core
COPY json ./json
COPY engine ./engine
COPY architecture-tests ./architecture-tests
COPY examples ./examples

RUN chmod +x gradlew
RUN --mount=type=cache,target=/root/.gradle,sharing=locked set --; \
    if [ -n "${GRIDGRIND_VERSION}" ]; then set -- "-Pversion=${GRIDGRIND_VERSION}"; fi; \
    ./gradlew --no-daemon :cli:shadowJar "$@"

# The literal FROM pin is checked against release/container-runtime.properties.
FROM debian:trixie-slim@sha256:89c1ea91e897f0fe0a3731595330bc45b8bdef4047464a07116f891a94054b31 AS runtime-base

COPY release/container-runtime.properties /usr/share/gridgrind/container-runtime.properties

# Snapshot metadata can expire; package authentication still uses Debian's archive keyring.
# HTTP bootstraps ca-certificates without trusting an additional bootstrap certificate source.
RUN set -eu; \
    . /usr/share/gridgrind/container-runtime.properties; \
    printf 'Types: deb deb-src\nURIs: http://snapshot.debian.org/archive/debian/%s/\nSuites: trixie trixie-updates\nComponents: main\nSigned-By: /usr/share/keyrings/debian-archive-keyring.pgp\nCheck-Valid-Until: no\n\nTypes: deb deb-src\nURIs: http://snapshot.debian.org/archive/debian-security/%s/\nSuites: trixie-security\nComponents: main\nSigned-By: /usr/share/keyrings/debian-archive-keyring.pgp\nCheck-Valid-Until: no\n' \
        "$DEBIAN_SNAPSHOT" "$DEBIAN_SNAPSHOT" > /etc/apt/sources.list.d/debian.sources

FROM runtime-base AS java-runtime

ARG TARGETARCH

# The JDK archive and download tools stay in this stage. Keep the complete vendor module
# inventory while omitting duplicated jmods, headers and manuals from the runtime image.
RUN set -eu; \
    . /usr/share/gridgrind/container-runtime.properties; \
    export DEBIAN_FRONTEND=noninteractive; \
    apt-get update >/dev/null; \
    apt-get install --no-install-recommends -y binutils ca-certificates curl libstdc++6 zlib1g >/dev/null; \
    case "$TARGETARCH" in \
        amd64) runtime_url=$CORRETTO_AMD64_URL; runtime_sha256=$CORRETTO_AMD64_SHA256 ;; \
        arm64) runtime_url=$CORRETTO_ARM64_URL; runtime_sha256=$CORRETTO_ARM64_SHA256 ;; \
        *) printf 'Unsupported runtime architecture: %s\n' "$TARGETARCH" >&2; exit 1 ;; \
    esac; \
    curl --fail --silent --show-error --location --retry 2 --connect-timeout 15 --max-time 240 \
        --output /tmp/corretto.tar.gz "$runtime_url"; \
    printf '%s  /tmp/corretto.tar.gz\n' "$runtime_sha256" | sha256sum --check --status; \
    mkdir /tmp/corretto; \
    tar -xzf /tmp/corretto.tar.gz -C /tmp/corretto --strip-components=1; \
    test "$(cat /tmp/corretto/commitId.txt)" = "$CORRETTO_SOURCE_COMMIT"; \
    test "$(cat /tmp/corretto/version.txt)" = "$CORRETTO_VERSION"; \
    /tmp/corretto/bin/jlink --module-path /tmp/corretto/jmods \
        --add-modules ALL-MODULE-PATH --strip-debug --no-header-files --no-man-pages \
        --compress=zip-6 --output /opt/java/openjdk; \
    /tmp/corretto/bin/java --list-modules | sort > /tmp/vendor-modules; \
    /opt/java/openjdk/bin/java --list-modules | sort > /tmp/runtime-modules; \
    cmp /tmp/vendor-modules /tmp/runtime-modules; \
    cp -a /tmp/corretto/legal/. /opt/java/openjdk/legal/; \
    cp /tmp/corretto/LICENSE /tmp/corretto/ADDITIONAL_LICENSE_INFO \
        /tmp/corretto/ASSEMBLY_EXCEPTION /tmp/corretto/README.md \
        /tmp/corretto/commitId.txt /tmp/corretto/version.txt /tmp/corretto/release \
        /opt/java/openjdk/; \
    rm -rf /tmp/corretto /tmp/corretto.tar.gz /tmp/vendor-modules /tmp/runtime-modules \
        /var/lib/apt/lists/*

FROM runtime-base

ARG GRIDGRIND_VERSION=unknown
ARG GRIDGRIND_REVISION=unknown

LABEL org.opencontainers.image.title="GridGrind"
LABEL org.opencontainers.image.description=".xlsx workbook automation from a JSON request"
LABEL org.opencontainers.image.version="${GRIDGRIND_VERSION}"
LABEL org.opencontainers.image.revision="${GRIDGRIND_REVISION}"
LABEL org.opencontainers.image.vendor="Ervins Strauhmanis"
LABEL org.opencontainers.image.source="https://github.com/resoltico/GridGrind"
LABEL org.opencontainers.image.documentation="https://github.com/resoltico/GridGrind/blob/main/README.md"
LABEL org.opencontainers.image.base.name="docker.io/library/debian:trixie-slim"
LABEL org.opencontainers.image.source-info="https://github.com/resoltico/GridGrind/releases/tag/v${GRIDGRIND_VERSION}"
LABEL org.opencontainers.image.corresponding-source="ghcr.io/resoltico/gridgrind:${GRIDGRIND_VERSION}-sources"

ARG GRIDGRIND_UID=65532
ARG GRIDGRIND_GID=65532

# Signature-line preview generation relies on Java2D/fontconfig even in headless mode.
# Ship a minimal deterministic font stack and route HOME/XDG cache into tmp-backed directories so
# the container stays quiet even when callers override --user to match host file ownership.
RUN export DEBIAN_FRONTEND=noninteractive \
    && apt-get update >/dev/null \
    && apt-get install --no-install-recommends -y ca-certificates fontconfig fonts-dejavu-core \
        libasound2t64 libstdc++6 zlib1g libx11-6 libxext6 libxi6 libxrender1 libxtst6 >/dev/null \
    && sed -i 's|http://snapshot.debian.org/|https://snapshot.debian.org/|g' /etc/apt/sources.list.d/debian.sources \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd --system --gid "${GRIDGRIND_GID}" gridgrind \
    && useradd --system --no-create-home --uid "${GRIDGRIND_UID}" --gid gridgrind --home-dir /home/gridgrind gridgrind \
    && install -d -o gridgrind -g gridgrind /home/gridgrind /work \
    && install -d -m 1777 /tmp/gridgrind-home /tmp/gridgrind-cache /tmp/gridgrind-cache/fontconfig \
    && fc-cache -f >/dev/null
COPY --from=java-runtime /opt/java/openjdk /opt/java/openjdk
ENV JAVA_HOME=/opt/java/openjdk
ENV PATH="/opt/java/openjdk/bin:${PATH}"
ENV HOME=/tmp/gridgrind-home
ENV XDG_CACHE_HOME=/tmp/gridgrind-cache

WORKDIR /work

COPY --from=build --chown=gridgrind:gridgrind /workspace/cli/build/libs/gridgrind.jar /app/gridgrind.jar

# Legal files are also embedded in META-INF/ inside gridgrind.jar.
# Copying them to the image filesystem makes them discoverable without unpacking the JAR.
COPY LICENSE /usr/share/doc/gridgrind/LICENSE
COPY NOTICE /usr/share/doc/gridgrind/NOTICE
COPY PATENTS.md /usr/share/doc/gridgrind/PATENTS.md
COPY LICENSE-APACHE-2.0 /usr/share/doc/gridgrind/LICENSE-APACHE-2.0
COPY LICENSE-BSD-2-CLAUSE /usr/share/doc/gridgrind/LICENSE-BSD-2-CLAUSE
COPY LICENSE-BSD-3-CLAUSE /usr/share/doc/gridgrind/LICENSE-BSD-3-CLAUSE

USER gridgrind:gridgrind

# Arguments select an explicit GridGrind subcommand, such as help or run.
ENTRYPOINT ["java", "-jar", "/app/gridgrind.jar"]
