package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Tests for local file hyperlink normalization and workbook-relative resolution. */
class ExcelFileHyperlinkTargetsTest {
  @Test
  void normalizePathAcceptsPlainPathsAndWindowsDrivePaths() {
    assertEquals("/tmp/report.xlsx", ExcelFileHyperlinkTargets.normalizePath("/tmp/report.xlsx"));
    assertEquals(
        "C:/temp/report.xlsx", ExcelFileHyperlinkTargets.normalizePath("C:/temp/report.xlsx"));
    assertEquals(
        "C:\\temp\\report.xlsx", ExcelFileHyperlinkTargets.normalizePath("C:\\temp\\report.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelFileHyperlinkTargets.normalizePath("C:temp/report.xlsx"));
  }

  @Test
  void toPoiAddressEscapesRelativeAndAbsolutePathsForPoi() {
    assertEquals(
        "support/budget%20backup.xlsx",
        ExcelFileHyperlinkTargets.toPoiAddress("support/budget backup.xlsx"));
    assertEquals(
        "reports/report-1_2~3.xlsx",
        ExcelFileHyperlinkTargets.toPoiAddress("reports/report-1_2~3.xlsx"));
    assertEquals("reports/%25zz.xlsx", ExcelFileHyperlinkTargets.toPoiAddress("reports/%zz.xlsx"));
    assertEquals(
        Path.of("/tmp/report file.xlsx").toUri().toASCIIString(),
        ExcelFileHyperlinkTargets.toPoiAddress("/tmp/report file.xlsx"));
    assertEquals(
        "file:///C:/temp/report%20file.xlsx",
        ExcelFileHyperlinkTargets.toPoiAddress("C:/temp/report file.xlsx"));
  }

  @Test
  void toPoiAddressRejectsPathsThatAreInvalidForTheLocalRuntime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelFileHyperlinkTargets.toPoiAddress("\u0000report.xlsx"));
    assertFalse(failure.getMessage().isBlank());
  }

  @Test
  void normalizePathNormalizesFileUrisWithAuthorities() {
    assertEquals(
        "//server/share/report.xlsx",
        ExcelFileHyperlinkTargets.normalizePath("file://server/share/report.xlsx"));
  }

  @Test
  void normalizePathDecodesEscapedRelativePoiPaths() {
    assertEquals(
        "support/budget backup.xlsx",
        ExcelFileHyperlinkTargets.normalizePath("support/budget%20backup.xlsx"));
    assertEquals(
        "reports/report#1.xlsx",
        ExcelFileHyperlinkTargets.normalizePath("reports/report%231.xlsx"));
    assertEquals(
        "C:/temp/%20report.xlsx",
        ExcelFileHyperlinkTargets.normalizePath("C:/temp/%20report.xlsx"));
    assertEquals("%20", ExcelFileHyperlinkTargets.normalizePath("%20"));
    assertEquals("reports/%zz.xlsx", ExcelFileHyperlinkTargets.normalizePath("reports/%zz.xlsx"));
  }

  @Test
  void normalizePathRejectsMissingAndMalformedFileUris() {
    IllegalArgumentException missingPath =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelFileHyperlinkTargets.normalizePath("file://server"));
    assertEquals("path must contain a file-system path", missingPath.getMessage());
    assertNotNull(missingPath.getCause());

    IllegalArgumentException malformedUri =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelFileHyperlinkTargets.normalizePath("file://%zz"));
    assertEquals("path must be a valid file: URI", malformedUri.getMessage());
    assertNotNull(malformedUri.getCause());

    IllegalArgumentException opaqueFileUriWithoutPath =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelFileHyperlinkTargets.normalizePath("file:relative-report.xlsx"));
    assertEquals("path must contain a file-system path", opaqueFileUriWithoutPath.getMessage());
    assertNotNull(opaqueFileUriWithoutPath.getCause());
  }

  @Test
  void normalizePathTreatsMalformedAbsoluteUriSyntaxAsPlainPathText() {
    assertEquals(
        "https://exa mple.com/report",
        ExcelFileHyperlinkTargets.normalizePath("https://exa mple.com/report"));
    assertEquals("C", ExcelFileHyperlinkTargets.normalizePath("C"));
  }

  @Test
  void resolveReturnsMalformedForInvalidLocalPaths() {
    FileHyperlinkResolution resolution =
        ExcelFileHyperlinkTargets.resolve(
            "\u0000report.xlsx",
            new WorkbookLocation.StoredWorkbook(Path.of("tmp", "Budget.xlsx").toAbsolutePath()));

    FileHyperlinkResolution.MalformedPath malformedPath =
        assertInstanceOf(FileHyperlinkResolution.MalformedPath.class, resolution);
    assertEquals("\u0000report.xlsx", malformedPath.path());
    assertFalse(malformedPath.reason().isBlank());
  }

  @Test
  void resolveReturnsMalformedForInvalidFileUriFallbackPaths() {
    FileHyperlinkResolution resolution =
        ExcelFileHyperlinkTargets.resolve(
            "file:///%00report.xlsx",
            new WorkbookLocation.StoredWorkbook(Path.of("tmp", "Budget.xlsx").toAbsolutePath()));

    FileHyperlinkResolution.MalformedPath malformedPath =
        assertInstanceOf(FileHyperlinkResolution.MalformedPath.class, resolution);
    assertEquals("/\u0000report.xlsx", malformedPath.path());
    assertFalse(malformedPath.reason().isBlank());
  }

  @Test
  void resolveAcceptsPoiEscapedRelativePaths() {
    FileHyperlinkResolution resolution =
        ExcelFileHyperlinkTargets.resolve(
            "reports/%zz.xlsx",
            new WorkbookLocation.StoredWorkbook(Path.of("tmp", "Budget.xlsx").toAbsolutePath()));

    FileHyperlinkResolution.ResolvedPath resolvedPath =
        assertInstanceOf(FileHyperlinkResolution.ResolvedPath.class, resolution);
    assertEquals("reports/%zz.xlsx", resolvedPath.path());
    assertTrue(resolvedPath.resolvedPath().endsWith(Path.of("tmp", "reports", "%zz.xlsx")));
  }

  @Test
  void resolveReturnsUnresolvedRelativePathForUnsavedWorkbooks() {
    FileHyperlinkResolution resolution =
        ExcelFileHyperlinkTargets.resolve(
            "reports/q1.xlsx", new WorkbookLocation.UnsavedWorkbook());

    FileHyperlinkResolution.UnresolvedRelativePath unresolved =
        assertInstanceOf(FileHyperlinkResolution.UnresolvedRelativePath.class, resolution);
    assertEquals("reports/q1.xlsx", unresolved.path());
  }
}
