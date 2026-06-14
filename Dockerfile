# syntax=docker/dockerfile:1.7

# Pin both builder and runtime manifest lists so local rebuilds and published images stay
# reproducible across time.
FROM azul/zulu-openjdk-alpine:26@sha256:33b52f3e06d325140b85bc67ddaf4731ca640b76bc1f15b78ddb292b56d9d8bf AS build

WORKDIR /workspace

COPY gradlew gradle.properties settings.gradle.kts build.gradle.kts ./
COPY gradle ./gradle
COPY authoring-java ./authoring-java
COPY cli ./cli
COPY contract ./contract
COPY engine ./engine
COPY excel-foundation ./excel-foundation
COPY executor ./executor

RUN chmod +x gradlew
RUN --mount=type=cache,target=/root/.gradle ./gradlew --no-daemon :cli:shadowJar

FROM azul/zulu-openjdk-alpine:26-jre@sha256:4202b612ef7e434db932b8c23d4d97c7bbd2b5a2d86be4e934c2676f2ee6bb57

LABEL org.opencontainers.image.licenses="MIT AND Apache-2.0 AND BSD-2-Clause AND BSD-3-Clause AND EDL-1.0"
LABEL org.opencontainers.image.vendor="Ervins Strauhmanis"

ARG GRIDGRIND_UID=65532
ARG GRIDGRIND_GID=65532

# Signature-line preview generation relies on Java2D/fontconfig even in headless mode.
# Ship a minimal deterministic font stack and route HOME/XDG cache into tmp-backed directories so
# the container stays quiet even when callers override --user to match host file ownership.
RUN apk add --no-cache fontconfig ttf-dejavu >/dev/null \
    && addgroup -g "${GRIDGRIND_GID}" -S gridgrind \
    && adduser -S -D -H -u "${GRIDGRIND_UID}" -G gridgrind -h /home/gridgrind gridgrind \
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
COPY LICENSE-EDL-1.0 /usr/share/doc/gridgrind/LICENSE-EDL-1.0

USER gridgrind:gridgrind

ENTRYPOINT ["java", "-jar", "/app/gridgrind.jar"]
