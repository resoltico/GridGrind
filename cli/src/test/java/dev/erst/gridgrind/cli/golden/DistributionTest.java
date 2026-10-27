package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.HashSet;
import java.util.HexFormat;
import java.util.List;
import java.util.Set;
import java.util.jar.JarFile;
import java.util.zip.ZipFile;
import org.apache.commons.compress.archivers.tar.TarArchiveInputStream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Checks the actual delivery archives, inherited legal material and installed launcher. */
class DistributionTest {
  private static final List<String> LEGAL =
      List.of(
          "LICENSE",
          "NOTICE",
          "PATENTS.md",
          "LICENSE-APACHE-2.0",
          "LICENSE-BSD-2-CLAUSE",
          "LICENSE-BSD-3-CLAUSE");
  @TempDir Path temporary;

  @Test
  void jarRetainsItsEntryPointLegalTextsAndAuditedFixture() throws Exception {
    verifyJar(GoldenLauncher.REPOSITORY.resolve("cli/build/libs/gridgrind.jar"));
  }

  @SuppressWarnings(
      "PMD.AvoidInstantiatingObjectsInLoops") // Decode each retained legal text independently.
  static void verifyJar(Path path) throws IOException, java.security.NoSuchAlgorithmException {
    try (JarFile jar = new JarFile(path.toFile())) {
      var attributes = jar.getManifest().getMainAttributes();
      assertEquals("dev.erst.gridgrind.cli.App", attributes.getValue("Main-Class"));
      assertEquals("Multiple; see META-INF/NOTICE", attributes.getValue("Implementation-License"));
      for (String name : LEGAL) {
        assertEquals(
            Files.readString(GoldenLauncher.REPOSITORY.resolve(name)),
            new String(
                jar.getInputStream(jar.getJarEntry("META-INF/" + name)).readAllBytes(),
                java.nio.charset.StandardCharsets.UTF_8));
      }
      for (String name :
          List.of(
              "FastDoubleParser-LICENSE",
              "FastDoubleParser-ThirdParty-LICENSE",
              "Schubfach-LICENSE")) {
        assertTrue(jar.getJarEntry("META-INF/" + name) != null, name);
      }
      assertFalse(jar.stream().anyMatch(entry -> entry.getName().contains("LICENSE-EDL-1.0")));
      byte[] fixture =
          jar.getInputStream(
                  jar.getJarEntry(
                      "gridgrind/recipe-assets/custom-xml-assets/custom-xml-mapping.xlsx"))
              .readAllBytes();
      assertEquals(
          "6a5fbb160c7c9c2add5125ac07ceaf3f841d40e03d5c75025ab71946d49af289",
          HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(fixture)));
      Set<String> names = new HashSet<>();
      var entries = jar.entries();
      while (entries.hasMoreElements())
        assertTrue(names.add(entries.nextElement().getName()), "Duplicate JAR entry");
    }
  }

  @Test
  void zipAndTarContainTheSameRunnableDistribution() throws IOException {
    Path directory = GoldenLauncher.REPOSITORY.resolve("cli/build/distributions");
    Set<String> zipNames = new HashSet<>();
    try (ZipFile archive = new ZipFile(singleArchive(directory, ".zip").toFile())) {
      archive.stream()
          .filter(entry -> !entry.isDirectory())
          .forEach(entry -> zipNames.add(entry.getName()));
    }
    Set<String> tarNames = new HashSet<>();
    try (TarArchiveInputStream archive =
        new TarArchiveInputStream(Files.newInputStream(singleArchive(directory, ".tar")))) {
      for (var entry = archive.getNextEntry(); entry != null; entry = archive.getNextEntry()) {
        if (!entry.isDirectory()) tarNames.add(entry.getName());
      }
    }
    assertEquals(zipNames, tarNames);
    String prefix =
        zipNames.stream()
            .filter(name -> name.endsWith("/bin/gridgrind"))
            .findFirst()
            .orElseThrow()
            .replace("bin/gridgrind", "");
    assertTrue(zipNames.contains(prefix + "bin/gridgrind.bat"));
    assertTrue(zipNames.contains(prefix + "lib/gridgrind.jar"));
    for (String name : LEGAL) assertTrue(zipNames.contains(prefix + name), name);
    assertEquals(
        1,
        zipNames.stream().filter(name -> name.contains("/lib/") && name.endsWith(".jar")).count());
  }

  @Test
  void installedLauncherPreservesPayloadAndExistingResponseOnPublicationFailure() throws Exception {
    Path installed = GoldenLauncher.REPOSITORY.resolve("cli/build/install/gridgrind");
    for (String name : LEGAL) assertTrue(Files.isRegularFile(installed.resolve(name)), name);
    String launcher = installed.resolve("bin/gridgrind").toString();
    var version = GoldenLauncher.runProcess(List.of(launcher, "version"), temporary, 60);
    assertEquals(0, version.exit(), version.stderr());
    assertTrue(version.stdout().startsWith("GridGrind "));
    assertEquals("", version.stderr());
    Path existing = temporary.resolve("existing response.json");
    Files.writeString(existing, "sentinel\n");
    var failed =
        GoldenLauncher.runProcess(
            List.of(launcher, "version", "--response", existing.toString()), temporary, 60);
    assertEquals(1, failed.exit());
    assertEquals("sentinel\n", Files.readString(existing));
    assertEquals(version.stdout(), failed.stdout());
    var notice = GoldenSnapshots.JSON.readTree(failed.stderr());
    assertEquals(3, notice.size());
    assertEquals("RESPONSE_WRITE_FAILED", notice.path("reason").asText());
    assertEquals("STDOUT", notice.path("wroteTo").asText());
    assertEquals(existing.toString(), notice.path("responsePath").asText());
    var license = GoldenLauncher.runProcess(List.of(launcher, "license"), temporary, 60);
    assertEquals(0, license.exit(), license.stderr());
    for (String text :
        List.of(
            "Mozilla Public License",
            "Jakarta Activation API",
            "Jakarta XML Binding API",
            "Eclipse Distribution License v1.0",
            "Boost Software License, Version 1.0")) {
      assertTrue(license.stdout().contains(text), text);
    }
  }

  @Test
  void installedLauncherExecutesAndReopensAPublishedWorkbook() throws Exception {
    String executable =
        GoldenLauncher.REPOSITORY.resolve("cli/build/install/gridgrind/bin/gridgrind").toString();
    WorkbookLauncher launcher =
        (root, arguments) -> {
          List<String> invocation = new java.util.ArrayList<>(List.of(executable));
          invocation.addAll(List.of(arguments));
          return GoldenLauncher.runProcess(invocation, root, 60);
        };
    Path published = temporary.resolve("published starter");
    var materialized =
        launcher.run(
            temporary, "recipe", "materialize", "BUDGET", "--workspace", published.toString());
    assertEquals(0, materialized.exit(), materialized.stderr());
    GoldenWorkbook.execute(launcher, published, published.resolve("budget-request.json"));
  }

  private static Path singleArchive(Path directory, String suffix) throws IOException {
    Path archive =
        directory.resolve(
            "gridgrind-" + dev.erst.gridgrind.cli.render.ProductInfo.version() + suffix);
    assertTrue(Files.isRegularFile(archive), () -> "Missing current distribution: " + archive);
    return archive;
  }
}
