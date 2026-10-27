# syntax=docker/dockerfile:1.28.0@sha256:bb22d9815c728170f72750f4e5b0d672e06176142e1d602c7e66c050100b7e5b

# Pin both builder and runtime manifest lists so local rebuilds and published images stay
# reproducible across time.
FROM azul-zulu:27.30-27.0.1-jdk-debian13@sha256:7f3f17771a7a8500b8844c7da21b89f14feb856b7bafe6d68292bca8e746f035 AS build

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
RUN --mount=type=cache,target=/root/.gradle set --; \
    if [ -n "${GRIDGRIND_VERSION}" ]; then set -- "-Pversion=${GRIDGRIND_VERSION}"; fi; \
    ./gradlew --no-daemon :cli:shadowJar "$@"

FROM azul-zulu:27.30-27.0.1-jre-debian13@sha256:5682efdef44a68c0b27f7a8a7365945dc662b209af6dfca9983d65db2a478394

ARG GRIDGRIND_VERSION=unknown

LABEL org.opencontainers.image.title="GridGrind"
LABEL org.opencontainers.image.description=".xlsx workbook automation from a JSON request"
LABEL org.opencontainers.image.version="${GRIDGRIND_VERSION}"
LABEL org.opencontainers.image.vendor="Ervins Strauhmanis"
LABEL org.opencontainers.image.source="https://github.com/resoltico/GridGrind"
LABEL org.opencontainers.image.documentation="https://github.com/resoltico/GridGrind/blob/main/README.md"
LABEL org.opencontainers.image.base.name="docker.io/library/azul-zulu:27.30-27.0.1-jre-debian13"

ARG GRIDGRIND_UID=65532
ARG GRIDGRIND_GID=65532

# Signature-line preview generation relies on Java2D/fontconfig even in headless mode.
# Ship a minimal deterministic font stack and route HOME/XDG cache into tmp-backed directories so
# the container stays quiet even when callers override --user to match host file ownership.
RUN apt-get update >/dev/null \
    && apt-get install --no-install-recommends -y fontconfig fonts-dejavu-core >/dev/null \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd --system --gid "${GRIDGRIND_GID}" gridgrind \
    && useradd --system --no-create-home --uid "${GRIDGRIND_UID}" --gid gridgrind --home-dir /home/gridgrind gridgrind \
    && install -d -o gridgrind -g gridgrind /home/gridgrind /work \
    && install -d -m 1777 /tmp/gridgrind-home /tmp/gridgrind-cache /tmp/gridgrind-cache/fontconfig \
    && fc-cache -f >/dev/null
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
