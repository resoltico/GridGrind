FROM eclipse-temurin:26-jre-alpine

WORKDIR /app

# The fat JAR is built by the GitHub Actions workflow (./gradlew :cli:shadowJar)
# before docker build is invoked. For local use, run that command first.
COPY cli/build/libs/*.jar gridgrind.jar

ENTRYPOINT ["java", "-jar", "gridgrind.jar"]
