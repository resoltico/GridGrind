# Pin the multi-arch base-image manifest list digest so rebuilds stay reproducible across time.
FROM azul/zulu-openjdk-alpine:26-jre@sha256:ba82b19503943a38b01a7fc77d4d25b87ae46d030cdd56aaca3a449d502c8179

LABEL org.opencontainers.image.licenses="MIT AND Apache-2.0 AND BSD-3-Clause"
LABEL org.opencontainers.image.vendor="Ervins Strauhmanis"

# Signature-line preview generation relies on Java2D/fontconfig even in headless mode.
# Ship a minimal deterministic font stack and a writable cache location so Docker matches the
# fat-JAR surface even when requests run as an arbitrary mounted-path UID:GID.
RUN apk add --no-cache fontconfig ttf-dejavu >/dev/null \
    && fc-cache -f >/dev/null
ENV HOME=/tmp
ENV XDG_CACHE_HOME=/tmp/.cache
RUN install -d -m 1777 /tmp/.cache /tmp/.cache/fontconfig

WORKDIR /app

# The fat JAR is built by the GitHub Actions workflow (./gradlew :cli:shadowJar)
# before docker build is invoked. For local use, run that command first.
COPY cli/build/libs/gridgrind.jar gridgrind.jar

# Legal files are also embedded in META-INF/ inside gridgrind.jar.
# Copying them to the image filesystem makes them discoverable without unpacking the JAR.
COPY LICENSE /usr/share/doc/gridgrind/LICENSE
COPY NOTICE /usr/share/doc/gridgrind/NOTICE
COPY PATENTS.md /usr/share/doc/gridgrind/PATENTS.md
COPY LICENSE-APACHE-2.0 /usr/share/doc/gridgrind/LICENSE-APACHE-2.0
COPY LICENSE-BSD-3-CLAUSE /usr/share/doc/gridgrind/LICENSE-BSD-3-CLAUSE

ENTRYPOINT ["java", "-jar", "/app/gridgrind.jar"]
