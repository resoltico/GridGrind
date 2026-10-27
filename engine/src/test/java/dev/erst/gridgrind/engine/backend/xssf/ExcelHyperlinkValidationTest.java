package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.engine.links.LinkAdapter;
import org.junit.jupiter.api.Test;

/** Tests for non-throwing hyperlink validation helpers. */
class ExcelHyperlinkValidationTest {
  @Test
  void validatesHyperlinkTargetsAcrossAllSupportedForms() {
    assertTrue(LinkAdapter.isValidUrlTarget("https://example.com/report"));
    assertFalse(LinkAdapter.isValidUrlTarget(null));
    assertFalse(LinkAdapter.isValidUrlTarget(" "));
    assertFalse(LinkAdapter.isValidUrlTarget("example.com/report"));
    assertFalse(LinkAdapter.isValidUrlTarget("https://exa mple.com/report"));
    assertFalse(LinkAdapter.isValidUrlTarget("https://[invalid"));

    assertTrue(LinkAdapter.isValidEmailTarget("team@example.com"));
    assertTrue(LinkAdapter.isValidEmailTarget("mailto:team@example.com"));
    assertTrue(LinkAdapter.isValidEmailTarget("MAILTO:team@example.com"));
    assertFalse(LinkAdapter.isValidEmailTarget(null));
    assertFalse(LinkAdapter.isValidEmailTarget("mailto:"));
    assertFalse(LinkAdapter.isValidEmailTarget("notanemail"));
    assertFalse(LinkAdapter.isValidEmailTarget("@example.com"));
    assertFalse(LinkAdapter.isValidEmailTarget("user@"));
    assertFalse(LinkAdapter.isValidEmailTarget("team @example.com"));
    assertFalse(LinkAdapter.isValidEmailTarget("team@@example.com"));

    assertTrue(LinkAdapter.isValidFileTarget("/tmp/report.xlsx"));
    assertFalse(LinkAdapter.isValidFileTarget(null));
    assertFalse(LinkAdapter.isValidFileTarget(" "));

    assertTrue(LinkAdapter.isValidDocumentTarget("Budget!A1"));
    assertFalse(LinkAdapter.isValidDocumentTarget(null));
    assertFalse(LinkAdapter.isValidDocumentTarget(" "));
  }

  @Test
  void stripsMailtoPrefixCaseInsensitively() {
    assertEquals("team@example.com", LinkAdapter.stripMailtoPrefix("mailto:team@example.com"));
    assertEquals("team@example.com", LinkAdapter.stripMailtoPrefix("MAILTO:team@example.com"));
    assertEquals("team@example.com", LinkAdapter.stripMailtoPrefix("team@example.com"));
  }

  @Test
  void absoluteUriSchemeReturnsNullForBlankTargetsAndTheSchemeForAbsoluteUris() {
    assertEquals(java.util.Optional.empty(), LinkAdapter.absoluteUriScheme(null));
    assertEquals(java.util.Optional.empty(), LinkAdapter.absoluteUriScheme(" "));
    assertEquals(
        java.util.Optional.of("https"),
        LinkAdapter.absoluteUriScheme("https://example.com/report"));
  }
}
