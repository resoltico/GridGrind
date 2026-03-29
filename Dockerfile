FROM azul/zulu-openjdk-alpine:26-jre

WORKDIR /app

# The fat JAR is built by the GitHub Actions workflow (./gradlew :cli:shadowJar)
# before docker build is invoked. For local use, run that command first.
COPY cli/build/libs/gridgrind.jar gridgrind.jar

ENTRYPOINT ["java", "-jar", "gridgrind.jar"]
