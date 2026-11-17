# syntax=docker/dockerfile:1.29.0@sha256:572cebbc2b15e6f84c077742952674889cb63cd5e9b80abdffcbe587b6f19183

# All Java roles share this immutable JDK/Ubuntu base. The pin is checked against
# release/container-runtime.properties by the image/source contract tests.
FROM eclipse-temurin:27.0.1.0-jdk-resolute@sha256:653c788ead947fce2b3452b434f25f69adf9b60c9f41e4639b3592caa719f683 AS java-base

COPY release/container-runtime.properties /usr/share/gridgrind/container-runtime.properties

# Fixed Ubuntu indexes remain authenticated by the upstream archive keyring.
# Upgrade inherited packages against the same snapshot used for additions and source collection.
RUN set -eu; \
    . /usr/share/gridgrind/container-runtime.properties; \
    rm -f /etc/apt/sources.list /etc/apt/sources.list.d/*.list /etc/apt/sources.list.d/*.sources; \
    printf 'Types: deb deb-src\nURIs: https://snapshot.ubuntu.com/ubuntu/%s/\nSuites: resolute resolute-updates resolute-security\nComponents: main universe\nSigned-By: /usr/share/keyrings/ubuntu-archive-keyring.gpg\nCheck-Valid-Until: no\n' \
        "$UBUNTU_SNAPSHOT" > /etc/apt/sources.list.d/ubuntu.sources; \
    export DEBIAN_FRONTEND=noninteractive; \
    apt-get update >/dev/null; \
    apt-get upgrade --no-install-recommends -y >/dev/null; \
    apt-get install --no-install-recommends -y fonts-dejavu-core >/dev/null; \
    fc-cache -f >/dev/null; \
    rm -rf /var/lib/apt/lists/*

# Default locale is explicit because formatted values are observable workbook facts.
# Keep Temurin's installed truststore; its certificate-mutating shell entrypoint is
# bypassed for both role targets, so USE_SYSTEM_CA_CERTS is not a product option.
ENV LANG=en_US.UTF-8 LANGUAGE=en_US:en LC_ALL=en_US.UTF-8 TZ=UTC
ENTRYPOINT []
CMD []

FROM java-base AS contributor
SHELL ["/bin/bash", "-o", "pipefail", "-c"]
RUN export DEBIAN_FRONTEND=noninteractive \
    && apt-get update >/dev/null \
    && apt-get install -y --no-install-recommends \
        ca-certificates curl git gh gnupg jq less lsof procps python3 shellcheck sudo unzip zip >/dev/null \
    && if getent passwd 1000 >/dev/null; then userdel --remove "$(getent passwd 1000 | cut -d: -f1)"; fi \
    && useradd --create-home --uid 1000 --shell /bin/bash vscode \
    && printf 'vscode ALL=(ALL) NOPASSWD:ALL\n' > /etc/sudoers.d/vscode \
    && chmod 0440 /etc/sudoers.d/vscode \
    && rm -rf /var/lib/apt/lists/*
RUN set -eu; \
    . /usr/share/gridgrind/container-runtime.properties; \
    case "$(dpkg --print-architecture)" in \
        amd64) buildx_sha256=$BUILDX_AMD64_SHA256; docker_cli_url=$DOCKER_CLI_AMD64_URL; docker_cli_sha256=$DOCKER_CLI_AMD64_SHA256 ;; \
        arm64) buildx_sha256=$BUILDX_ARM64_SHA256; docker_cli_url=$DOCKER_CLI_ARM64_URL; docker_cli_sha256=$DOCKER_CLI_ARM64_SHA256 ;; \
        *) exit 1 ;; \
    esac; \
    curl -fLsS "$docker_cli_url" -o /tmp/docker-cli.deb; \
    printf '%s  /tmp/docker-cli.deb\n' "$docker_cli_sha256" | sha256sum --check --status; \
    dpkg --install /tmp/docker-cli.deb; \
    test "$(docker --version | cut -d' ' -f3 | tr -d ',')" = "$DOCKER_CLI_VERSION"; \
    rm /tmp/docker-cli.deb; \
    install -d /usr/local/lib/docker/cli-plugins; \
    curl -fLsS "https://github.com/docker/buildx/releases/download/v${BUILDX_VERSION}/buildx-v${BUILDX_VERSION}.linux-$(dpkg --print-architecture)" \
        -o /usr/local/lib/docker/cli-plugins/docker-buildx; \
    printf '%s  /usr/local/lib/docker/cli-plugins/docker-buildx\n' "$buildx_sha256" | sha256sum --check --status; \
    chmod 0755 /usr/local/lib/docker/cli-plugins/docker-buildx
USER vscode
ENV HOME=/home/vscode
WORKDIR /workspaces/gridgrind

FROM java-base AS build

ARG GRIDGRIND_VERSION

WORKDIR /workspace

COPY gradlew gradle.properties settings.gradle.kts build.gradle.kts gradle.lockfile settings-gradle.lockfile ./
COPY gradle ./gradle
COPY release/container-runtime.properties ./release/container-runtime.properties
COPY LICENSE PROJECT-NOTICE.md LICENSE-COMMERCIAL.md NOTICE PATENTS.md LICENSE-APACHE-2.0 LICENSE-BSD-2-CLAUSE LICENSE-BSD-3-CLAUSE ./
COPY cli ./cli
COPY core ./core
COPY json ./json
COPY engine ./engine
COPY native-provider ./native-provider
COPY architecture-tests ./architecture-tests
COPY examples ./examples

RUN chmod +x gradlew
RUN --mount=type=cache,target=/root/.gradle,sharing=locked set --; \
    if [ -n "${GRIDGRIND_VERSION}" ]; then set -- "-Pversion=${GRIDGRIND_VERSION}"; fi; \
    ./gradlew --no-daemon :cli:shadowJar "$@"

FROM java-base AS application

ARG GRIDGRIND_VERSION=unknown
ARG GRIDGRIND_REVISION=unknown

LABEL org.opencontainers.image.title="GridGrind"
LABEL org.opencontainers.image.description=".xlsx workbook automation from a JSON request"
LABEL org.opencontainers.image.version="${GRIDGRIND_VERSION}"
LABEL org.opencontainers.image.revision="${GRIDGRIND_REVISION}"
LABEL org.opencontainers.image.vendor="Ervins Strauhmanis"
LABEL org.opencontainers.image.source="https://github.com/resoltico/GridGrind"
LABEL org.opencontainers.image.documentation="https://github.com/resoltico/GridGrind/blob/main/README.md"
LABEL org.opencontainers.image.base.name="docker.io/library/eclipse-temurin:27.0.1.0-jdk-resolute"
LABEL org.opencontainers.image.source-info="https://github.com/resoltico/GridGrind/releases/tag/v${GRIDGRIND_VERSION}"
LABEL org.opencontainers.image.corresponding-source="ghcr.io/resoltico/gridgrind:${GRIDGRIND_VERSION}-sources"

ARG GRIDGRIND_UID=65532
ARG GRIDGRIND_GID=65532

RUN groupadd --system --gid "${GRIDGRIND_GID}" gridgrind \
    && useradd --system --no-create-home --uid "${GRIDGRIND_UID}" --gid gridgrind --home-dir /home/gridgrind gridgrind \
    && install -d -o gridgrind -g gridgrind /home/gridgrind /work \
    && install -d -m 1777 /tmp/gridgrind-home /tmp/gridgrind-cache /tmp/gridgrind-cache/fontconfig
ENV HOME=/tmp/gridgrind-home
ENV XDG_CACHE_HOME=/tmp/gridgrind-cache

WORKDIR /work

COPY --from=build --chown=gridgrind:gridgrind /workspace/cli/build/libs/gridgrind.jar /app/gridgrind.jar

# Legal files are also embedded in META-INF/ inside gridgrind.jar.
# Copying them to the image filesystem makes them discoverable without unpacking the JAR.
COPY LICENSE /usr/share/doc/gridgrind/LICENSE
COPY PROJECT-NOTICE.md /usr/share/doc/gridgrind/PROJECT-NOTICE.md
COPY LICENSE-COMMERCIAL.md /usr/share/doc/gridgrind/LICENSE-COMMERCIAL.md
COPY NOTICE /usr/share/doc/gridgrind/NOTICE
COPY PATENTS.md /usr/share/doc/gridgrind/PATENTS.md
COPY LICENSE-APACHE-2.0 /usr/share/doc/gridgrind/LICENSE-APACHE-2.0
COPY LICENSE-BSD-2-CLAUSE /usr/share/doc/gridgrind/LICENSE-BSD-2-CLAUSE
COPY LICENSE-BSD-3-CLAUSE /usr/share/doc/gridgrind/LICENSE-BSD-3-CLAUSE

USER gridgrind:gridgrind

# Arguments select an explicit GridGrind subcommand, such as help or run.
ENTRYPOINT ["java", "-jar", "/app/gridgrind.jar"]
