# syntax=docker/dockerfile:1.7

# Pin both builder and runtime manifest lists so local rebuilds and published images stay
# reproducible across time.
FROM azul/zulu-openjdk:26@sha256:456ddce6098187ea8b9069cbf141b6a124d1fdf667818c195ba95be6a0e48e70 AS build

WORKDIR /workspace

COPY gradlew gradle.properties settings.gradle.kts build.gradle.kts ./
COPY gradle ./gradle
COPY LICENSE NOTICE PATENTS.md LICENSE-APACHE-2.0 LICENSE-BSD-2-CLAUSE LICENSE-BSD-3-CLAUSE ./
COPY authoring-java ./authoring-java
COPY cli ./cli
COPY contract ./contract
COPY engine ./engine
COPY excel-foundation ./excel-foundation
COPY executor ./executor
COPY examples ./examples

RUN chmod +x gradlew
RUN --mount=type=cache,target=/root/.gradle ./gradlew --no-daemon :cli:shadowJar

FROM azul/zulu-openjdk:26-jre@sha256:ac36910df585bf3db5a38b30695eb04791515d1bb7d78202564db560c60c3470

ARG GRIDGRIND_VERSION=unknown

LABEL org.opencontainers.image.title="GridGrind"
LABEL org.opencontainers.image.description=".xlsx workbook automation from a JSON request"
LABEL org.opencontainers.image.version="${GRIDGRIND_VERSION}"
LABEL org.opencontainers.image.vendor="Ervins Strauhmanis"
LABEL org.opencontainers.image.source="https://github.com/resoltico/GridGrind"
LABEL org.opencontainers.image.documentation="https://github.com/resoltico/GridGrind/blob/main/README.md"
LABEL org.opencontainers.image.base.name="docker.io/azul/zulu-openjdk:26-jre"

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

ENTRYPOINT ["java", "-jar", "/app/gridgrind.jar"]
