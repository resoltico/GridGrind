package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.microsoft.schemas.office.spreadsheetml.x2018.threadedcomments.CTThreadedComments;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.CodeSource;
import java.util.List;
import java.util.Locale;
import java.util.jar.JarEntry;
import java.util.jar.JarFile;
import java.util.stream.Stream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Build-failing audit over the Apache POI runtime and full-schema XLSX capability surface. */
class ApachePoiRuntimeCapabilityAuditTest {
  @Test
  void runtimeJarIncludesDocumentedFirstClassApis() {
    assertAll(
        () ->
            assertDoesNotThrow(
                () ->
                    Class.forName(
                        "org.apache.poi.xssf.usermodel.XSSFSignatureLine",
                        false,
                        contextClassLoader()),
                "poi-ooxml runtime must ship XSSFSignatureLine"),
        () ->
            assertDoesNotThrow(
                () ->
                    Class.forName(
                        "org.apache.poi.xssf.extractor.XSSFImportFromXML",
                        false,
                        contextClassLoader()),
                "poi-ooxml runtime must ship XSSFImportFromXML"),
        () ->
            assertDoesNotThrow(
                () ->
                    Class.forName(
                        "org.apache.poi.xssf.extractor.XSSFExportToXml",
                        false,
                        contextClassLoader()),
                "poi-ooxml runtime must ship XSSFExportToXml"),
        () ->
            assertDoesNotThrow(
                () ->
                    Class.forName("org.apache.poi.xssf.model.MapInfo", false, contextClassLoader()),
                "poi-ooxml runtime must ship MapInfo"));
  }

  @Test
  void runtimeAndFullSchemaJarsExplainRemainingUnexposedFamilies() throws Exception {
    List<String> runtimeThreaded = matchingEntries(XSSFWorkbook.class, "threaded");
    List<String> runtimeSpark = matchingEntries(XSSFWorkbook.class, "spark");
    List<String> runtimeSlicer = matchingEntries(XSSFWorkbook.class, "slicer");
    List<String> fullThreaded = matchingEntries(CTThreadedComments.class, "threaded");
    List<String> fullSpark = matchingEntries(CTThreadedComments.class, "spark");
    List<String> fullSlicer = matchingEntries(CTThreadedComments.class, "slicer");

    assertAll(
        () ->
            assertTrue(
                runtimeThreaded.isEmpty(),
                () ->
                    "poi-ooxml runtime unexpectedly exposes threaded entries: " + runtimeThreaded),
        () ->
            assertTrue(
                runtimeSpark.isEmpty(),
                () -> "poi-ooxml runtime unexpectedly exposes sparkline entries: " + runtimeSpark),
        () ->
            assertTrue(
                runtimeSlicer.isEmpty(),
                () -> "poi-ooxml runtime unexpectedly exposes slicer entries: " + runtimeSlicer),
        () ->
            assertFalse(
                fullThreaded.isEmpty(), "poi-ooxml-full must expose threaded-comment schema types"),
        () ->
            assertTrue(
                fullThreaded.stream()
                    .anyMatch(entry -> entry.toLowerCase(Locale.ROOT).contains("threadedcomments")),
                () ->
                    "poi-ooxml-full threaded-comment schema entries missing expected naming: "
                        + fullThreaded),
        () ->
            assertTrue(
                fullSpark.isEmpty(),
                () -> "poi-ooxml-full unexpectedly exposes sparkline schema entries: " + fullSpark),
        () ->
            assertTrue(
                fullSlicer.isEmpty(),
                () -> "poi-ooxml-full unexpectedly exposes slicer schema entries: " + fullSlicer));
  }

  private static List<String> matchingEntries(Class<?> anchor, String needle)
      throws IOException, URISyntaxException {
    Path jarPath = jarPath(anchor);
    try (JarFile jarFile = new JarFile(jarPath.toFile())) {
      String loweredNeedle = needle.toLowerCase(Locale.ROOT);
      try (Stream<JarEntry> entries = jarFile.stream()) {
        return entries
            .map(JarEntry::getName)
            .filter(name -> name.toLowerCase(Locale.ROOT).contains(loweredNeedle))
            .sorted()
            .toList();
      }
    }
  }

  private static Path jarPath(Class<?> anchor) throws URISyntaxException {
    CodeSource codeSource = anchor.getProtectionDomain().getCodeSource();
    if (codeSource == null) {
      throw new IllegalStateException("No code source available for " + anchor.getName());
    }
    Path path = Path.of(codeSource.getLocation().toURI());
    if (!Files.isRegularFile(path)) {
      throw new IllegalStateException(
          "Expected dependency jar for " + anchor.getName() + " but found " + path);
    }
    return path;
  }

  private static ClassLoader contextClassLoader() {
    return Thread.currentThread().getContextClassLoader();
  }
}
