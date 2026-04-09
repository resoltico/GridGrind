package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Tests for the shared Excel sheet-name validation contract. */
class ExcelSheetNamesTest {
  @Test
  void acceptsValidExcelSheetNames() {
    assertDoesNotThrow(() -> ExcelSheetNames.requireValid("Budget 2026", "sheetName"));
  }

  @Test
  void rejectsReservedCharactersAndEdgeApostrophes() {
    IllegalArgumentException invalidCharacter =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelSheetNames.requireValid("Bad:Name", "sheetName"));
    assertTrue(invalidCharacter.getMessage().contains("invalid Excel character ':'"));

    IllegalArgumentException leadingQuote =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelSheetNames.requireValid("'Budget", "sheetName"));
    assertTrue(leadingQuote.getMessage().contains("single quote"));

    IllegalArgumentException trailingQuote =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelSheetNames.requireValid("Budget'", "sheetName"));
    assertTrue(trailingQuote.getMessage().contains("single quote"));
  }

  @Test
  void rejectsControlCharactersWithCodePointDetail() {
    IllegalArgumentException controlCharacter =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelSheetNames.requireValid("\u0003Budget", "sheetName"));

    assertTrue(controlCharacter.getMessage().contains("U+0003"));
  }
}
