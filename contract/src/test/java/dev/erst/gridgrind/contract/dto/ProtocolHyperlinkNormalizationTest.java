package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for protocol-owned hyperlink normalization helpers. */
class ProtocolHyperlinkNormalizationTest {
  @Test
  void normalizesAndRejectsUrlTargets() {
    assertEquals(
        "https://example.com/report?q=1",
        ProtocolHyperlinkUrlSupport.normalizeUrlTarget("https://example.com/report?q=1"));
    assertEquals(
        "http://example.com/page",
        ProtocolHyperlinkUrlSupport.normalizeUrlTarget("http://example.com/page"));
    assertEquals(
        "ftp://files.example.com/report.xlsx",
        ProtocolHyperlinkUrlSupport.normalizeUrlTarget("ftp://files.example.com/report.xlsx"));
    assertThrows(
        NullPointerException.class, () -> ProtocolHyperlinkUrlSupport.normalizeUrlTarget(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeUrlTarget("relative/path"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeUrlTarget("file:///tmp/report.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeUrlTarget("mailto:team@example.com"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeUrlTarget("http://[broken"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeUrlTarget("javascript:alert(1)"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeUrlTarget("vbscript:msgbox(1)"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ProtocolHyperlinkUrlSupport.normalizeUrlTarget(
                "ms-excel:ofe|u|https://example.com/evil.xlsx"));
    IllegalArgumentException schemeError =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ProtocolHyperlinkUrlSupport.normalizeUrlTarget("ldap://evil.example.com/payload"));
    assertTrue(schemeError.getMessage().contains("ldap"));
  }

  @Test
  void normalizesEmailAndDocumentTargets() {
    assertEquals(
        "team@example.com",
        ProtocolHyperlinkUrlSupport.normalizeEmailTarget("mailto:team@example.com"));
    assertEquals(
        "team@example.com", ProtocolHyperlinkUrlSupport.normalizeEmailTarget("team@example.com"));
    assertEquals("Budget!B4", ProtocolHyperlinkTargetSupport.normalizeDocumentTarget("Budget!B4"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeEmailTarget("team example.com"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkUrlSupport.normalizeEmailTarget("team@@example.com"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkTargetSupport.normalizeDocumentTarget(" "));
  }

  @Test
  void normalizesFileTargetsFromUrisAndEscapedPaths() {
    assertEquals("A", ProtocolHyperlinkFileSupport.normalizeFileTarget("A"));
    assertEquals(
        "/tmp/report.xlsx",
        ProtocolHyperlinkFileSupport.normalizeFileTarget("file:///tmp/report.xlsx"));
    assertEquals(
        "//server/share/report.xlsx",
        ProtocolHyperlinkFileSupport.normalizeFileTarget("file://server/share/report.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkFileSupport.normalizeFileTarget("file://server"));
    assertEquals(
        "team folder/report.xlsx",
        ProtocolHyperlinkFileSupport.normalizeFileTarget("team%20folder/report.xlsx"));
    assertEquals("?q=%20", ProtocolHyperlinkFileSupport.normalizeFileTarget("?q=%20"));
    assertEquals(
        "C:\\temp\\report.xlsx",
        ProtocolHyperlinkFileSupport.normalizeFileTarget("C:\\temp\\report.xlsx"));
    assertEquals(
        "C:/temp/report.xlsx",
        ProtocolHyperlinkFileSupport.normalizeFileTarget("C:/temp/report.xlsx"));
    assertThrows(
        NullPointerException.class, () -> ProtocolHyperlinkFileSupport.normalizeFileTarget(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkFileSupport.normalizeFileTarget(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkFileSupport.normalizeFileTarget("https://example.com/report.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkFileSupport.normalizeFileTarget("C:temp"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkFileSupport.normalizeFileTarget("file:///tmp/%00bad.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkFileSupport.normalizeFileTarget("file://"));
    assertEquals("bad%2", ProtocolHyperlinkFileSupport.normalizeFileTarget("bad%2"));
  }

  @Test
  void buildsPoiFileAddressesForRelativeAbsoluteAndWindowsPaths() {
    assertEquals("A", ProtocolHyperlinkFileSupport.toPoiFileAddress("A"));
    assertEquals(
        "Ab1/team%20folder/report.xlsx",
        ProtocolHyperlinkFileSupport.toPoiFileAddress("Ab1/team folder/report.xlsx"));
    assertEquals(
        "file:///tmp/report.xlsx",
        ProtocolHyperlinkFileSupport.toPoiFileAddress("/tmp/report.xlsx"));
    assertEquals(
        "file:///C:/temp/report.xlsx",
        ProtocolHyperlinkFileSupport.toPoiFileAddress("C:\\temp\\report.xlsx"));

    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> ProtocolHyperlinkFileSupport.toPoiFileAddress("\u0000bad.xlsx"));
    assertNotNull(exception.getMessage());
    assertFalse(exception.getMessage().isBlank());
  }
}
