package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for non-throwing hyperlink validation helpers. */
class ExcelHyperlinkValidationTest {
  @Test
  void validatesHyperlinkTargetsAcrossAllSupportedForms() {
    assertTrue(ExcelHyperlinkValidation.isValidUrlTarget("https://example.com/report"));
    assertFalse(ExcelHyperlinkValidation.isValidUrlTarget(null));
    assertFalse(ExcelHyperlinkValidation.isValidUrlTarget(" "));
    assertFalse(ExcelHyperlinkValidation.isValidUrlTarget("example.com/report"));
    assertFalse(ExcelHyperlinkValidation.isValidUrlTarget("https://exa mple.com/report"));
    assertFalse(ExcelHyperlinkValidation.isValidUrlTarget("https://[invalid"));

    assertTrue(ExcelHyperlinkValidation.isValidEmailTarget("team@example.com"));
    assertTrue(ExcelHyperlinkValidation.isValidEmailTarget("mailto:team@example.com"));
    assertTrue(ExcelHyperlinkValidation.isValidEmailTarget("MAILTO:team@example.com"));
    assertFalse(ExcelHyperlinkValidation.isValidEmailTarget(null));
    assertFalse(ExcelHyperlinkValidation.isValidEmailTarget("mailto:"));

    assertTrue(ExcelHyperlinkValidation.isValidFileTarget("/tmp/report.xlsx"));
    assertFalse(ExcelHyperlinkValidation.isValidFileTarget(null));
    assertFalse(ExcelHyperlinkValidation.isValidFileTarget(" "));

    assertTrue(ExcelHyperlinkValidation.isValidDocumentTarget("Budget!A1"));
    assertFalse(ExcelHyperlinkValidation.isValidDocumentTarget(null));
    assertFalse(ExcelHyperlinkValidation.isValidDocumentTarget(" "));
  }

  @Test
  void stripsMailtoPrefixCaseInsensitively() {
    assertEquals(
        "team@example.com", ExcelHyperlinkValidation.stripMailtoPrefix("mailto:team@example.com"));
    assertEquals(
        "team@example.com", ExcelHyperlinkValidation.stripMailtoPrefix("MAILTO:team@example.com"));
    assertEquals(
        "team@example.com", ExcelHyperlinkValidation.stripMailtoPrefix("team@example.com"));
  }
}
