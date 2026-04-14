FROM azul/zulu-openjdk-alpine:26-jre

LABEL org.opencontainers.image.licenses="MIT AND Apache-2.0 AND BSD-3-Clause"
LABEL org.opencontainers.image.vendor="Ervins Strauhmanis"

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
