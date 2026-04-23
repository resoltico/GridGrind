package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.Enumeration;
import java.util.Map;
import java.util.function.UnaryOperator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import org.junit.jupiter.api.Test;

/** Regression tests for persisted OOXML picture-package invariants. */
class XlsxPicturePackageInvariantSupportTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void requireCanonicalPicturePackageStateAcceptsCanonicalWorkbook() throws IOException {
    Path workbookPath = workbookWithPicture("gridgrind-picture-package-canonical-");

    assertDoesNotThrow(
        () -> XlsxPicturePackageInvariantSupport.requireCanonicalPicturePackageState(workbookPath));
  }

  @Test
  void requireCanonicalPicturePackageStateRejectsMissingPictureRelationships() throws IOException {
    Path workbookPath = workbookWithPicture("gridgrind-picture-package-missing-");
    Path mutatedWorkbook =
        rewriteEntry(
            workbookPath,
            "xl/drawings/drawing1.xml",
            xml -> xml.replaceFirst("r:embed=\"[^\"]+\"", "r:embed=\"rIdMissing\""));

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                XlsxPicturePackageInvariantSupport.requireCanonicalPicturePackageState(
                    mutatedWorkbook));
    assertTrue(failure.getMessage().contains("picture refs must resolve"));
    assertTrue(failure.getMessage().contains("rIdMissing"));
  }

  @Test
  void requireCanonicalPicturePackageStateAcceptsCopiedEmbeddedObjectPreviewRelations()
      throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-picture-package-embedded-copy-", ".xlsx");
    Files.deleteIfExists(workbookPath);
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .sheet("Ops")
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  new ExcelDrawingAnchor.TwoCell(
                      new ExcelDrawingMarker(1, 1, 0, 0),
                      new ExcelDrawingMarker(4, 6, 0, 0),
                      null)));
      workbook.copySheet("Ops", "Ops Copy", new ExcelSheetCopyPosition.AppendAtEnd());
      workbook.save(workbookPath);
    }

    assertDoesNotThrow(
        () -> XlsxPicturePackageInvariantSupport.requireCanonicalPicturePackageState(workbookPath));
  }

  @Test
  void requireCanonicalPicturePackageStateRejectsWorksheetPreviewRefsPointingAtDrawings()
      throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-picture-package-sheet-preview-", ".xlsx");
    Files.deleteIfExists(workbookPath);
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .sheet("Ops")
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  new ExcelDrawingAnchor.TwoCell(
                      new ExcelDrawingMarker(1, 1, 0, 0),
                      new ExcelDrawingMarker(4, 6, 0, 0),
                      null)));
      workbook.save(workbookPath);
    }

    Path mutatedWorkbook =
        rewriteEntry(
            workbookPath,
            "xl/worksheets/sheet1.xml",
            XlsxPicturePackageInvariantSupportTest::retargetObjectPreviewToSheetDrawing);

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                XlsxPicturePackageInvariantSupport.requireCanonicalPicturePackageState(
                    mutatedWorkbook));
    assertTrue(
        failure.getMessage().contains("embedded object preview refs must target /xl/media parts"));
    assertTrue(failure.getMessage().contains("xl/drawings/"));
  }

  private static Path workbookWithPicture(String prefix) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    Files.deleteIfExists(workbookPath);
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .sheet("Ops")
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(PNG_PIXEL_BYTES),
                  ExcelPictureFormat.PNG,
                  new ExcelDrawingAnchor.TwoCell(
                      new ExcelDrawingMarker(1, 1, 0, 0), new ExcelDrawingMarker(4, 6, 0, 0), null),
                  "Queue preview"));
      workbook.save(workbookPath);
    }
    return workbookPath;
  }

  private static Path rewriteEntry(
      Path sourceWorkbook, String entryName, UnaryOperator<String> transformer) throws IOException {
    return rewriteEntries(sourceWorkbook, Map.of(entryName, transformer));
  }

  @SuppressWarnings("PMD.AvoidInstantiatingObjectsInLoops")
  private static Path rewriteEntries(
      Path sourceWorkbook, Map<String, UnaryOperator<String>> transformers) throws IOException {
    Path mutatedWorkbook = Files.createTempFile("gridgrind-picture-package-mutated-", ".xlsx");
    try (ZipFile zipFile = new ZipFile(sourceWorkbook.toFile());
        ZipOutputStream outputStream =
            new ZipOutputStream(Files.newOutputStream(mutatedWorkbook))) {
      Enumeration<? extends ZipEntry> entries = zipFile.entries();
      while (entries.hasMoreElements()) {
        ZipEntry entry = entries.nextElement();
        try (InputStream inputStream = zipFile.getInputStream(entry)) {
          byte[] bytes = inputStream.readAllBytes();
          outputStream.putNextEntry(new ZipEntry(entry.getName()));
          UnaryOperator<String> transformer = transformers.get(entry.getName());
          outputStream.write(
              transformer == null
                  ? bytes
                  : transformer
                      .apply(new String(bytes, java.nio.charset.StandardCharsets.UTF_8))
                      .getBytes(java.nio.charset.StandardCharsets.UTF_8));
          outputStream.closeEntry();
        }
      }
    }
    return mutatedWorkbook;
  }

  private static String retargetObjectPreviewToSheetDrawing(String worksheetXml) {
    Matcher drawingMatcher =
        Pattern.compile("<(?:\\w+:)?drawing[^>]*\\br:id=\"([^\"]+)\"").matcher(worksheetXml);
    if (!drawingMatcher.find()) {
      throw new IllegalStateException("worksheet XML must declare a drawing relationship");
    }
    String drawingRelationId = drawingMatcher.group(1);
    Matcher objectPreviewMatcher =
        Pattern.compile("(<(?:\\w+:)?objectPr[^>]*\\br:id=\")([^\"]+)(\")").matcher(worksheetXml);
    if (!objectPreviewMatcher.find()) {
      throw new IllegalStateException(
          "worksheet XML must declare an embedded-object preview relationship");
    }
    return objectPreviewMatcher.replaceFirst("$1" + drawingRelationId + "$3");
  }
}
