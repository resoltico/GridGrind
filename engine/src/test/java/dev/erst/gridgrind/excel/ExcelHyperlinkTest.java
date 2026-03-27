package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelHyperlink record construction and normalization. */
class ExcelHyperlinkTest {
  @Test
  void buildsSupportedHyperlinkVariants() {
    assertEquals(ExcelHyperlinkType.URL, new ExcelHyperlink.Url("https://example.com").type());
    assertEquals(ExcelHyperlinkType.EMAIL, new ExcelHyperlink.Email("team@example.com").type());
    assertEquals("team@example.com", new ExcelHyperlink.Email("mailto:team@example.com").target());
    assertEquals("team@example.com", new ExcelHyperlink.Email("team@example.com").target());
    assertEquals(ExcelHyperlinkType.FILE, new ExcelHyperlink.File("/tmp/report.xlsx").type());
    assertEquals(ExcelHyperlinkType.DOCUMENT, new ExcelHyperlink.Document("Budget!B4").type());
  }

  @Test
  void validatesHyperlinkInputs() {
    assertThrows(NullPointerException.class, () -> new ExcelHyperlink.Url(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelHyperlink.Url("relative"));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelHyperlink.Url("https://exa mple.com"));
    assertThrows(NullPointerException.class, () -> new ExcelHyperlink.Email(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelHyperlink.Email("mailto:"));
    assertThrows(NullPointerException.class, () -> new ExcelHyperlink.File(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelHyperlink.File(" "));
    assertThrows(NullPointerException.class, () -> new ExcelHyperlink.Document(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelHyperlink.Document(" "));
  }
}
