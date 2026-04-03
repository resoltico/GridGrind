package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for protocol-owned hyperlink normalization helpers. */
class ProtocolHyperlinkSupportTest {
  @Test
  void normalizesAndRejectsUrlTargets() {
    assertEquals(
        "https://example.com/report?q=1",
        ProtocolHyperlinkSupport.normalizeUrlTarget("https://example.com/report?q=1"));
    assertThrows(
        NullPointerException.class, () -> ProtocolHyperlinkSupport.normalizeUrlTarget(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeUrlTarget("relative/path"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeUrlTarget("file:///tmp/report.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeUrlTarget("mailto:team@example.com"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeUrlTarget("http://[broken"));
  }

  @Test
  void normalizesEmailAndDocumentTargets() {
    assertEquals(
        "team@example.com",
        ProtocolHyperlinkSupport.normalizeEmailTarget("mailto:team@example.com"));
    assertEquals(
        "team@example.com", ProtocolHyperlinkSupport.normalizeEmailTarget("team@example.com"));
    assertEquals("Budget!B4", ProtocolHyperlinkSupport.normalizeDocumentTarget("Budget!B4"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeEmailTarget("team example.com"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeEmailTarget("team@@example.com"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeDocumentTarget(" "));
  }

  @Test
  void normalizesFileTargetsFromUrisAndEscapedPaths() {
    assertEquals("A", ProtocolHyperlinkSupport.normalizeFileTarget("A"));
    assertEquals(
        "/tmp/report.xlsx",
        ProtocolHyperlinkSupport.normalizeFileTarget("file:///tmp/report.xlsx"));
    assertEquals(
        "//server/share/report.xlsx",
        ProtocolHyperlinkSupport.normalizeFileTarget("file://server/share/report.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeFileTarget("file://server"));
    assertEquals(
        "team folder/report.xlsx",
        ProtocolHyperlinkSupport.normalizeFileTarget("team%20folder/report.xlsx"));
    assertEquals("?q=%20", ProtocolHyperlinkSupport.normalizeFileTarget("?q=%20"));
    assertEquals(
        "C:\\temp\\report.xlsx",
        ProtocolHyperlinkSupport.normalizeFileTarget("C:\\temp\\report.xlsx"));
    assertEquals(
        "C:/temp/report.xlsx", ProtocolHyperlinkSupport.normalizeFileTarget("C:/temp/report.xlsx"));
    assertThrows(
        NullPointerException.class, () -> ProtocolHyperlinkSupport.normalizeFileTarget(null));
    assertThrows(
        IllegalArgumentException.class, () -> ProtocolHyperlinkSupport.normalizeFileTarget(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeFileTarget("https://example.com/report.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeFileTarget("C:temp"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeFileTarget("file:///tmp/%00bad.xlsx"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolHyperlinkSupport.normalizeFileTarget("file://"));
    assertEquals("bad%2", ProtocolHyperlinkSupport.normalizeFileTarget("bad%2"));
  }

  @Test
  void buildsPoiFileAddressesForRelativeAbsoluteAndWindowsPaths() {
    assertEquals(
        "Ab1/team%20folder/report.xlsx",
        ProtocolHyperlinkSupport.toPoiFileAddress("Ab1/team folder/report.xlsx"));
    assertEquals(
        "file:///tmp/report.xlsx", ProtocolHyperlinkSupport.toPoiFileAddress("/tmp/report.xlsx"));
    assertEquals(
        "file:///C:/temp/report.xlsx",
        ProtocolHyperlinkSupport.toPoiFileAddress("C:\\temp\\report.xlsx"));

    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> ProtocolHyperlinkSupport.toPoiFileAddress("\u0000bad.xlsx"));
    assertNotNull(exception.getMessage());
    assertFalse(exception.getMessage().isBlank());
  }
}
