package dev.erst.gridgrind.cli.golden;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.DigestInputStream;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
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

  static void verifyJar(Path path) throws IOException, java.security.NoSuchAlgorithmException {
    try (JarFile jar = new JarFile(path.toFile())) {
      var attributes = jar.getManifest().getMainAttributes();
      assertEquals("dev.erst.gridgrind.cli.App", attributes.getValue("Main-Class"));
      assertEquals("Multiple; see META-INF/NOTICE", attributes.getValue("Implementation-License"));
      for (String name : LEGAL) {
        byte[] expected = Files.readAllBytes(GoldenLauncher.REPOSITORY.resolve(name));
        for (String prefix : List.of("META-INF/", "licenses/")) {
          assertArrayEquals(
              expected,
              jar.getInputStream(jar.getJarEntry(prefix + name)).readAllBytes(),
              prefix + name);
        }
      }
      for (String name :
          List.of(
              "FastDoubleParser-LICENSE",
              "FastDoubleParser-ThirdParty-LICENSE",
              "Schubfach-LICENSE")) {
        assertTrue(jar.getJarEntry("META-INF/" + name) != null, name);
      }
      for (String excluded :
          List.of(
              "com/ctc/wstx/shaded/msv_core/datatype/xsd/CommandLineTester.class",
              "META-INF/maven/com.sun.xml.bind.jaxb/isorelax/pom.xml")) {
        assertTrue(jar.getJarEntry(excluded) == null, excluded);
      }
      for (String retained :
          List.of(
              "META-INF/maven/com.sun.xml.bind.jaxb/isorelax/pom.properties",
              "com/ctc/wstx/shaded/msv/org_isorelax/dispatcher/SchemaProvider.class",
              "com/ctc/wstx/stax/WstxInputFactory.class",
              "com/ctc/wstx/stax/WstxOutputFactory.class",
              "com/ctc/wstx/stax/WstxEventFactory.class",
              "com/ctc/wstx/dtd/DTDSchemaFactory.class",
              "com/ctc/wstx/msv/RelaxNGSchemaFactory.class",
              "com/ctc/wstx/msv/W3CSchemaFactory.class",
              "com/ctc/wstx/shaded/msv_core/datatype/xsd/ngimpl/DataTypeLibraryImpl.class")) {
        assertTrue(jar.getJarEntry(retained) != null, retained);
      }
      for (String toolPackage :
          List.of(
              "edu/berkeley/cs/jqf/",
              "com/code_intelligence/jazzer/",
              "org/junit/",
              "org/pitest/",
              "net/sourceforge/pmd/",
              "dev/erst/gridgrind/fuzz/")) {
        assertFalse(
            jar.stream().anyMatch(entry -> entry.getName().startsWith(toolPackage)), toolPackage);
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
  @SuppressWarnings("PMD.UseConcurrentHashMap") // One invocation owns both sequential digest maps.
  void zipAndTarContainTheSameRunnableDistribution() throws IOException, NoSuchAlgorithmException {
    Path directory = GoldenLauncher.REPOSITORY.resolve("cli/build/distributions");
    Map<String, String> zipContents = new HashMap<>();
    try (ZipFile archive = new ZipFile(singleArchive(directory, ".zip").toFile())) {
      var entries = archive.entries();
      while (entries.hasMoreElements()) {
        var entry = entries.nextElement();
        if (!entry.isDirectory()) {
          try (InputStream input = archive.getInputStream(entry)) {
            assertTrue(zipContents.put(entry.getName(), sha256(input)) == null, entry.getName());
          }
        }
      }
    }
    Map<String, String> tarContents = new HashMap<>();
    try (TarArchiveInputStream archive =
        new TarArchiveInputStream(Files.newInputStream(singleArchive(directory, ".tar")))) {
      for (var entry = archive.getNextEntry(); entry != null; entry = archive.getNextEntry()) {
        if (!entry.isDirectory()) {
          assertTrue(tarContents.put(entry.getName(), sha256(archive)) == null, entry.getName());
        }
      }
    }
    assertEquals(zipContents, tarContents);
    Set<String> names = zipContents.keySet();
    String prefix =
        names.stream()
            .filter(name -> name.endsWith("/bin/gridgrind"))
            .findFirst()
            .orElseThrow()
            .replace("bin/gridgrind", "");
    assertTrue(names.contains(prefix + "bin/gridgrind.bat"));
    try (InputStream input =
        Files.newInputStream(GoldenLauncher.REPOSITORY.resolve("cli/build/libs/gridgrind.jar"))) {
      assertEquals(sha256(input), zipContents.get(prefix + "lib/gridgrind.jar"));
    }
    for (String name : LEGAL) {
      try (InputStream input = Files.newInputStream(GoldenLauncher.REPOSITORY.resolve(name))) {
        assertEquals(sha256(input), zipContents.get(prefix + name), name);
      }
    }
    assertEquals(
        1, names.stream().filter(name -> name.contains("/lib/") && name.endsWith(".jar")).count());
  }

  private static String sha256(InputStream input) throws IOException, NoSuchAlgorithmException {
    MessageDigest digest = MessageDigest.getInstance("SHA-256");
    // Do not close the stream here: TAR callers must advance to the next archive entry.
    new DigestInputStream(input, digest).transferTo(OutputStream.nullOutputStream());
    return HexFormat.of().formatHex(digest.digest());
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
