package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for selector helper validation and normalization seams. */
class SelectorSupportTest {
  @Test
  void selectorCardinalityFlagsMatchTheDeclaredContract() {
    assertAll(
        () -> assertFalse(SelectorCardinality.EXACTLY_ONE.allowsMany()),
        () -> assertFalse(SelectorCardinality.EXACTLY_ONE.allowsZero()),
        () -> assertFalse(SelectorCardinality.ZERO_OR_ONE.allowsMany()),
        () -> assertTrue(SelectorCardinality.ZERO_OR_ONE.allowsZero()),
        () -> assertTrue(SelectorCardinality.ONE_OR_MORE.allowsMany()),
        () -> assertFalse(SelectorCardinality.ONE_OR_MORE.allowsZero()),
        () -> assertTrue(SelectorCardinality.ANY_NUMBER.allowsMany()),
        () -> assertTrue(SelectorCardinality.ANY_NUMBER.allowsZero()));
  }

  @Test
  void selectorSupportValidatesScalarFieldsAndA1Geometry() {
    assertEquals("Budget", SelectorSupport.requireNonBlank("Budget", "field"));
    assertEquals("Budget", SelectorSupport.requireSheetName("Budget", "sheetName"));
    assertEquals("Budget_Total", SelectorSupport.requireDefinedName("Budget_Total", "selector"));
    assertEquals(
        "Sales Pivot 2026", SelectorSupport.requirePivotTableName("Sales Pivot 2026", "selector"));
    assertEquals("$B$12", SelectorSupport.requireAddress("$B$12", "selector"));
    assertEquals("A1:B3", SelectorSupport.requireRange("A1:B3", "selector"));
    assertEquals(4, SelectorSupport.requirePositive(4, "count"));
    assertEquals(0, SelectorSupport.requireNonNegative(0, "index"));
    assertEquals(-2, SelectorSupport.requireNonZero(-2, "delta"));
    assertEquals(1_048_576 - 1, SelectorSupport.requireRowIndexWithinBounds(1_048_576 - 1, "row"));
    assertEquals(16_384 - 1, SelectorSupport.requireColumnIndexWithinBounds(16_384 - 1, "column"));
    SelectorSupport.requireWindowSize(500, 500);
    assertEquals("AA10", SelectorSupport.absoluteA1Address(9, 26));
    assertEquals(26, SelectorSupport.columnIndex("$AA$10"));
    assertEquals(9, SelectorSupport.rowIndex("$AA$10"));
  }

  @Test
  void selectorSupportRejectsInvalidScalarFieldsAndA1Geometry() {
    assertEquals(
        "field must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> SelectorSupport.requireNonBlank(" ", "field"))
            .getMessage());
    assertEquals(
        "field must not be null",
        assertThrows(
                NullPointerException.class, () -> SelectorSupport.requireNonBlank(null, "field"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requireDefinedName("1bad", "selector"))
            .getMessage()
            .startsWith("selector "));
    assertEquals(
        "name must start with a letter or underscore and contain only letters, digits, underscore, or period",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requireDefinedName("1bad", "name"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requirePivotTableName(" ", "selector"))
            .getMessage()
            .startsWith("selector "));
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requireAddress("A0", "selector"))
            .getMessage()
            .startsWith("selector "));
    assertEquals(
        "selector must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> SelectorSupport.requireRange(" ", "selector"))
            .getMessage());
    assertEquals(
        "selector must be a rectangular A1-style range with at most one ':'",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requireRange("A1:B2:C3", "selector"))
            .getMessage());
    assertEquals(
        "count must be greater than 0",
        assertThrows(
                IllegalArgumentException.class, () -> SelectorSupport.requirePositive(0, "count"))
            .getMessage());
    assertEquals(
        "index must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requireNonNegative(-1, "index"))
            .getMessage());
    assertEquals(
        "delta must not be 0",
        assertThrows(
                IllegalArgumentException.class, () -> SelectorSupport.requireNonZero(0, "delta"))
            .getMessage());
    assertEquals(
        "row must be within Excel .xlsx row bounds",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requireRowIndexWithinBounds(1_048_576, "row"))
            .getMessage());
    assertEquals(
        "column must be within Excel .xlsx column bounds",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.requireColumnIndexWithinBounds(16_384, "column"))
            .getMessage());
    assertEquals(
        "rowCount * columnCount must not exceed 250000 but was 250500",
        assertThrows(
                IllegalArgumentException.class, () -> SelectorSupport.requireWindowSize(501, 500))
            .getMessage());
  }

  @Test
  void selectorSupportCopiesDistinctCollectionsAndRejectsDuplicateOrNullEntries() {
    assertEquals(
        List.of("A1", "B2"),
        SelectorSupport.copyDistinctAddresses(List.of("A1", "B2"), "addresses"));
    assertEquals(
        List.of("A1", "B2"), SelectorSupport.copyDistinctAddresses(List.of("A1", "B2"), "cells"));
    assertEquals(List.of("A1:B2"), SelectorSupport.copyDistinctRanges(List.of("A1:B2"), "ranges"));
    assertEquals(
        List.of("Budget", "Ops"),
        SelectorSupport.copyDistinctSheetNames(List.of("Budget", "Ops"), "sheetNames"));
    assertEquals(
        List.of("BudgetTotal"),
        SelectorSupport.copyDistinctDefinedNames(List.of("BudgetTotal"), "names"));
    assertEquals(
        List.of("Sales Pivot 2026"),
        SelectorSupport.copyDistinctPivotTableNames(List.of("Sales Pivot 2026"), "names"));
    assertEquals(
        List.of("x", "y"), SelectorSupport.copyDistinctValues(List.of("x", "y"), "values"));
    assertEquals(
        List.of(
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new NamedRangeSelector.SheetScope("LocalItem", "Budget")),
        SelectorSupport.copyDistinctNamedRangeRefs(
            List.of(
                new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                new NamedRangeSelector.SheetScope("LocalItem", "Budget")),
            "selectors"));

    assertEquals(
        "addresses must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.copyDistinctAddresses(List.of("A1", "A1"), "addresses"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.copyDistinctAddresses(List.of("A0"), "cells"))
            .getMessage()
            .startsWith("cells[0] address "));
    assertEquals(
        "ranges must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.copyDistinctRanges(List.of("A1:B2", "A1:B2"), "ranges"))
            .getMessage());
    assertEquals(
        "sheetNames must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorSupport.copyDistinctSheetNames(
                        List.of("Budget", "Budget"), "sheetNames"))
            .getMessage());
    assertEquals(
        "names must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorSupport.copyDistinctDefinedNames(
                        List.of("BudgetTotal", "BudgetTotal"), "names"))
            .getMessage());
    assertEquals(
        "names must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorSupport.copyDistinctPivotTableNames(
                        List.of("Sales Pivot 2026", "sales pivot 2026"), "names"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.copyDistinctPivotTableNames(List.of(" "), "names"))
            .getMessage()
            .startsWith("names[0] "));
    assertEquals(
        "values[1] must not be null",
        assertThrows(
                NullPointerException.class,
                () -> SelectorSupport.copyDistinctValues(Arrays.asList("x", null), "values"))
            .getMessage());
    assertEquals(
        "values must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.copyDistinctValues(List.of(), "values"))
            .getMessage());
    assertEquals(
        "values must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.copyDistinctValues(List.of("x", "x"), "values"))
            .getMessage());
    assertEquals(
        "selectors must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorSupport.copyDistinctNamedRangeRefs(List.of(), "selectors"))
            .getMessage());
    assertEquals(
        "selectors must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorSupport.copyDistinctNamedRangeRefs(
                        List.of(
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new NamedRangeSelector.WorkbookScope("budgettotal")),
                        "selectors"))
            .getMessage());
  }

  @Test
  void selectorSupportCoversZeroIterationAddressParsingAndCatalogLookupOrdering() {
    assertEquals(-1, SelectorSupport.columnIndex(""));
    assertEquals(-1, SelectorSupport.columnIndex("123"));
    assertThrows(NumberFormatException.class, () -> SelectorSupport.rowIndex(""));
    assertEquals(122, SelectorSupport.rowIndex("123"));
  }

  @Test
  void prefixedValidationMessagePreservesNullBlankAndAlreadyPrefixedMessages() {
    assertNull(SelectorSupport.prefixedValidationMessage("field", null));
    assertEquals(" ", SelectorSupport.prefixedValidationMessage("field", " "));
    assertEquals(
        "field must not be blank",
        SelectorSupport.prefixedValidationMessage("field", "field must not be blank"));
    assertEquals("field invalid", SelectorSupport.prefixedValidationMessage("field", "invalid"));
  }
}
