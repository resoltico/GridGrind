# Contributor image with the full JDK, shell tooling and fonts used for development.
# The shipped product image uses a separate JRE and carries only runtime components.
FROM mcr.microsoft.com/devcontainers/base:3.0.10-ubuntu26.04@sha256:91eff046d5c12318fc98e59258982b87ba896c467427ed8b5dab5a02107a5db7

SHELL ["/bin/bash", "-o", "pipefail", "-c"]

USER root

RUN apt-get update \
    && export DEBIAN_FRONTEND=noninteractive \
    && apt-get install -y --no-install-recommends \
        ca-certificates \
        curl \
        fontconfig \
        fonts-dejavu \
        gh \
        git \
        gnupg \
        jq \
        less \
        lsof \
        procps \
        python3 \
        shellcheck \
        unzip \
        zip \
    && curl -fsSL https://repos.azul.com/azul-repo.key \
        | gpg --dearmor -o /usr/share/keyrings/azul.gpg \
    && chmod 644 /usr/share/keyrings/azul.gpg \
    && echo "deb [signed-by=/usr/share/keyrings/azul.gpg] https://repos.azul.com/zulu/deb stable main" \
        > /etc/apt/sources.list.d/zulu.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends zulu27-jdk \
    && zulu_home="$(find /usr/lib/jvm -maxdepth 1 -mindepth 1 -type d -name 'zulu27*' | sort | head -1)" \
    && [[ -n "${zulu_home}" ]] \
    && ln -sfn "${zulu_home}" /usr/lib/jvm/zulu-27 \
    && test -x /usr/lib/jvm/zulu-27/bin/java \
    && test -x /usr/lib/jvm/zulu-27/bin/javac \
    && fc-cache -f >/dev/null \
    && java --version >/dev/null \
    && javac --version >/dev/null \
    && git --version >/dev/null \
    && shellcheck --version >/dev/null \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

ENV JAVA_HOME=/usr/lib/jvm/zulu-27
ENV PATH=${JAVA_HOME}/bin:${PATH}

USER vscode
