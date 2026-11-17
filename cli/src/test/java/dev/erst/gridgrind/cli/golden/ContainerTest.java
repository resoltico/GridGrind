package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.render.ProductInfo;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.io.TempDir;

/**
 * Actual image, legal, published-recipe and backend checks; explicitly selected by containerTest.
 */
@Tag("container")
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class ContainerTest {
  private Path workspace;
  private DockerFixture docker;

  @BeforeAll
  void prepareImage(@TempDir Path directory) throws Exception {
    workspace = directory;
    String prebuiltImage = System.getProperty("gridgrind.container.image", "");
    String requestedPlatform =
        DockerFixture.validateRequestedPlatform(
            System.getProperty("gridgrind.container.platform", ""));
    if (prebuiltImage.isBlank() && !requestedPlatform.isEmpty())
      throw new IllegalArgumentException("containerPlatform requires a borrowed containerImage");
    docker =
        prebuiltImage.isBlank()
            ? new DockerFixture(workspace)
            : new DockerFixture(workspace, prebuiltImage, requestedPlatform);
    verifyPinnedPlatforms();
    if (prebuiltImage.isBlank()) {
      docker.build(
          GoldenLauncher.REPOSITORY,
          List.of("--build-arg", "GRIDGRIND_VERSION=" + ProductInfo.version()));
    }
  }

  @AfterAll
  void removeOwnedResources() throws Exception {
    if (docker != null) docker.close();
  }

  @Test
  void packagedIdentityLegalFilesAndPinnedPlatformsArePresent() throws Exception {
    var help = docker.runWorkbook(workspace, "help");
    assertTrue(DockerFixture.assertSuccess(help).contains("Usage:"));
    assertEquals("", help.stderr());
    var version =
        GoldenSnapshots.JSON.readTree(
            DockerFixture.assertSuccess(
                docker.runWorkbook(workspace, "version", "--format", "structured")));
    assertEquals(ProductInfo.version(), version.path("version").asText());
    assertTrue(
        DockerFixture.assertSuccess(docker.runWorkbook(workspace, "license"))
            .contains("GNU AFFERO GENERAL PUBLIC LICENSE"));
    assertEquals(
        "GridGrind",
        docker.inspectImage("{{index .Config.Labels \"org.opencontainers.image.title\"}}").strip());
    String licenseLabel =
        docker
            .inspectImage("{{index .Config.Labels \"org.opencontainers.image.licenses\"}}")
            .strip();
    assertTrue(licenseLabel.isEmpty() || "<no value>".equals(licenseLabel));
    DockerFixture.assertSuccess(
        docker.run(
            List.of("--entrypoint", "sh"),
            "-c",
            "set -eu; for f in LICENSE NOTICE PATENTS.md PROJECT-NOTICE.md LICENSE-COMMERCIAL.md"
                + " LICENSE-APACHE-2.0 LICENSE-BSD-2-CLAUSE LICENSE-BSD-3-CLAUSE; do test -s"
                + " /usr/share/doc/gridgrind/$f || exit 1; done; test -f"
                + " /usr/share/doc/base-files/copyright; test -x \"$JAVA_HOME/bin/javac\"; test -s"
                + " \"$JAVA_HOME/NOTICE\"; test -s \"$JAVA_HOME/release\"; test -d"
                + " \"$JAVA_HOME/legal\"; find \"$JAVA_HOME/legal\" -type f -print -quit | grep -q"
                + " ."));
    var sha256 = java.security.MessageDigest.getInstance("SHA-256");
    var hexadecimal = java.util.HexFormat.of();
    var hashCommand = List.of("--entrypoint", "sha256sum");
    for (String name : List.of("PROJECT-NOTICE.md", "LICENSE-COMMERCIAL.md")) {
      String observed =
          DockerFixture.assertSuccess(docker.run(hashCommand, "/usr/share/doc/gridgrind/" + name));
      String expected =
          hexadecimal.formatHex(
              sha256.digest(Files.readAllBytes(GoldenLauncher.REPOSITORY.resolve(name))));
      assertEquals(expected, observed.substring(0, 64), name);
    }
    Properties runtime = new Properties();
    try (var input =
        Files.newInputStream(
            GoldenLauncher.REPOSITORY.resolve("release/container-runtime.properties"))) {
      runtime.load(input);
    }
    String identity =
        DockerFixture.assertSuccess(
            docker.run(List.of("--entrypoint", "sh"), "-c", "cat \"$JAVA_HOME/release\""));
    assertTrue(identity.contains("IMPLEMENTOR=\"Eclipse Adoptium\""));
    assertTrue(
        identity.contains(
            "JAVA_RUNTIME_VERSION=\"" + runtime.getProperty("TEMURIN_VERSION") + "\""));
    assertTrue(
        identity.contains(
            "BUILD_SOURCE=\"git:" + runtime.getProperty("TEMURIN_BUILD_COMMIT") + "\""));
    assertTrue(identity.contains(runtime.getProperty("TEMURIN_SOURCE_COMMIT").substring(0, 12)));
  }

  @Test
  void actualImageConfigurationAndJvmUseTheSelectedPlatform() throws Exception {
    String selected = docker.inspectedPlatform();
    var running =
        docker.run(List.of("--entrypoint", "java"), "-XshowSettings:properties", "-version");
    DockerFixture.assertSuccess(running);
    DockerFixture.requireJvmPlatform(selected, running.stdout() + running.stderr());
    String created =
        docker.create(List.of("--entrypoint", "java"), "-XshowSettings:properties", "-version");
    var started = docker.command(120, "start", "--attach", created);
    DockerFixture.assertSuccess(started);
    DockerFixture.requireJvmPlatform(selected, started.stdout() + started.stderr());
  }

  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  // One parser owns stage mappings; every external base needs its own platform set.
  private void verifyPinnedPlatforms() throws java.io.IOException, InterruptedException {
    Properties runtime = new Properties();
    try (var input =
        Files.newInputStream(
            GoldenLauncher.REPOSITORY.resolve("release/container-runtime.properties"))) {
      runtime.load(input);
    }
    Map<String, String> stages = new HashMap<>();
    String finalBase = "";
    for (String line : Files.readAllLines(GoldenLauncher.REPOSITORY.resolve("Dockerfile"))) {
      String[] words = line.strip().split("\\s+", -1);
      if (words.length < 2 || !"FROM".equalsIgnoreCase(words[0])) continue;
      int imageIndex = 1;
      while (imageIndex < words.length && words[imageIndex].startsWith("--")) imageIndex++;
      assertTrue(imageIndex < words.length, "FROM requires an image");
      String base = words[imageIndex];
      boolean localStage =
          stages.containsKey(base.toLowerCase(Locale.ROOT)) || "scratch".equals(base);
      finalBase = stages.getOrDefault(base.toLowerCase(Locale.ROOT), base);
      if (imageIndex + 2 < words.length && "AS".equalsIgnoreCase(words[imageIndex + 1]))
        stages.put(words[imageIndex + 2].toLowerCase(Locale.ROOT), finalBase);
      if (localStage) continue;
      var manifest =
          GoldenSnapshots.JSON.readTree(
              DockerFixture.assertSuccess(
                  docker.command(120, "buildx", "imagetools", "inspect", "--raw", base)));
      Set<String> platforms = new HashSet<>();
      for (var entry : manifest.path("manifests")) {
        platforms.add(
            entry.path("platform").path("os").asText()
                + "/"
                + entry.path("platform").path("architecture").asText());
      }
      assertTrue(platforms.containsAll(Set.of("linux/amd64", "linux/arm64")), platforms.toString());
    }
    assertEquals(
        runtime.getProperty("JAVA_BASE_IMAGE"),
        finalBase,
        "Final runtime base must match the source/packaging input pin");
  }

  @Test
  void everyPublishedRecipeMaterializesExecutesAndReopens() throws Exception {
    for (var recipe : RecipeRegistry.recipes()) {
      Path root = Files.createDirectories(workspace.resolve("recipes").resolve(recipe.id()));
      Path materialized = root.resolve("published workspace");
      DockerFixture.assertSuccess(
          docker.runWorkbook(
              root, "recipe", "materialize", recipe.id(), "--workspace", materialized.toString()));
      if ("PACKAGE_SECURITY_INSPECTION".equals(recipe.id())) {
        Path secrets = Files.createDirectories(materialized.resolve("fixture-secrets"));
        Files.writeString(secrets.resolve("source-open-password"), "GridGrind-2026");
      }
      GoldenWorkbook.execute(
          docker::runWorkbook, materialized, materialized.resolve(recipe.requestFileName()));
    }
  }

  @Test
  void customWorkdirProgressFontsStreamingAndReadbackWork() throws Exception {
    ContainerScenarios.verify(
        docker::runWorkbook, Files.createDirectories(workspace.resolve("smoke paths [#]")));
  }

  @Test
  void readOnlyRootSupportsHostUidWithExplicitWritableScratch() throws Exception {
    List<String> options = List.of("--read-only", "--tmpfs", "/tmp:rw,mode=1777");
    ContainerScenarios.verify(
        (root, arguments) -> docker.runWorkbook(options, root, arguments),
        Files.createDirectories(workspace.resolve("read-only-root")));
  }

  @Test
  void actualPkcs12SigningEncryptionAndDecryptionUseThePackagedProviders() throws Exception {
    Path root = Files.createDirectories(workspace.resolve("crypto"));
    Path secrets = Files.createDirectories(root.resolve("fixture-secrets"));
    Path password = secrets.resolve("crypto-password");
    Files.writeString(password, "Synthetic-GridGrind-Qualification-2026");
    Path keystore = root.resolve("signer.p12");
    DockerFixture.assertSuccess(
        docker.runWorkbook(
            List.of("--entrypoint", "keytool"),
            root,
            "-genkeypair",
            "-alias",
            "qualification",
            "-keyalg",
            "RSA",
            "-keysize",
            "2048",
            "-storetype",
            "PKCS12",
            "-keystore",
            keystore.toString(),
            "-storepass:file",
            password.toString(),
            "-dname",
            "CN=GridGrind synthetic test",
            "-validity",
            "2"));
    var plan =
        GoldenSnapshots.JSON.readTree(
            """
            {"protocolVersion":1,"source":{"type":"NEW"},
             "persistence":{"type":"SAVE_AS","path":"signed-encrypted.xlsx","ifExists":"REJECT",
              "security":{"encryption":{"type":"ENCRYPT","encryption":{"passwordRef":{"id":"crypto-password"}}},
               "signature":{"type":"SIGN","signature":{"pkcs12Path":"signer.p12",
                 "keystorePasswordRef":{"id":"crypto-password"}}}}},
             "steps":[{"stepId":"sheet","target":{"type":"SHEET_BY_NAME","name":"Verified"},
              "action":{"type":"ENSURE_SHEET"}}]}
            """);
    Path request = root.resolve("write.json");
    Files.writeString(request, GoldenSnapshots.JSON.writeValueAsString(plan));
    Path grant = GoldenWorkbook.grant(plan, root, "write-grant.json");
    var authority =
        (tools.jackson.databind.node.ObjectNode)
            GoldenSnapshots.JSON.readTree(Files.readAllBytes(grant));
    authority.put("maxInputSnapshotBytes", 1024 * 1024);
    ((tools.jackson.databind.node.ArrayNode) authority.path("readableResources"))
        .addObject()
        .put("type", "FILE")
        .put("path", keystore.toString());
    Files.writeString(grant, GoldenSnapshots.JSON.writeValueAsString(authority));
    var written =
        docker.runWorkbook(
            root,
            "run",
            "--request",
            request.toString(),
            "--grant",
            grant.toString(),
            "--secrets-provider",
            secrets.toString());
    assertEquals(
        "SUCCEEDED",
        GoldenSnapshots.JSON
            .readTree(DockerFixture.assertSuccess(written))
            .path("status")
            .asText());
    var readPlan =
        GoldenSnapshots.JSON.readTree(
            """
            {"protocolVersion":1,
             "source":{"type":"EXISTING","path":"signed-encrypted.xlsx",
                       "security":{"passwordRef":{"id":"crypto-password"}}},
             "persistence":{"type":"NONE"},
             "steps":[{"stepId":"security","target":{"type":"WORKBOOK_CURRENT"},
                       "query":{"type":"GET_PACKAGE_SECURITY"}}]}
            """);
    Path readRequest = root.resolve("read.json");
    Files.writeString(readRequest, GoldenSnapshots.JSON.writeValueAsString(readPlan));
    Path readGrant = GoldenWorkbook.grant(readPlan, root, "read-grant.json");
    var read =
        docker.runWorkbook(
            root,
            "run",
            "--request",
            readRequest.toString(),
            "--grant",
            readGrant.toString(),
            "--secrets-provider",
            secrets.toString());
    String result = DockerFixture.assertSuccess(read);
    var document = GoldenSnapshots.JSON.readTree(result);
    assertEquals("SUCCEEDED", document.path("status").asText());
    var security = document.path("inspections").get(0).path("security");
    assertEquals("ENCRYPTED", security.path("encryption").path("type").asText());
    assertEquals(1, security.path("signatures").size());
    assertEquals("VALID", security.path("signatures").get(0).path("state").asText());
  }
}
