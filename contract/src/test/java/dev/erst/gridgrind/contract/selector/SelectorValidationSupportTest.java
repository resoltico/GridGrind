package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for selector helper validation and normalization seams. */
class SelectorValidationSupportTest {
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
    assertEquals("Budget", SelectorValueValidation.requireNonBlank("Budget", "field"));
    assertEquals("Budget", SelectorValueValidation.requireSheetName("Budget", "sheetName"));
    assertEquals(
        "Budget_Total", SelectorValueValidation.requireDefinedName("Budget_Total", "selector"));
    assertEquals(
        "Sales Pivot 2026",
        SelectorValueValidation.requirePivotTableName("Sales Pivot 2026", "selector"));
    assertEquals("$B$12", SelectorValueValidation.requireAddress("$B$12", "selector"));
    assertEquals("A1:B3", SelectorValueValidation.requireRange("A1:B3", "selector"));
    assertEquals(4, SelectorValueValidation.requirePositive(4, "count"));
    assertEquals(0, SelectorValueValidation.requireNonNegative(0, "index"));
    assertEquals(-2, SelectorValueValidation.requireNonZero(-2, "delta"));
    assertEquals(
        1_048_576 - 1, SelectorValueValidation.requireRowIndexWithinBounds(1_048_576 - 1, "row"));
    assertEquals(
        16_384 - 1, SelectorValueValidation.requireColumnIndexWithinBounds(16_384 - 1, "column"));
    SelectorValueValidation.requireWindowSize(500, 500);
    assertEquals("AA10", SelectorAddressSupport.absoluteA1Address(9, 26));
    assertEquals(26, SelectorAddressSupport.columnIndex("$AA$10"));
    assertEquals(9, SelectorAddressSupport.rowIndex("$AA$10"));
  }

  @Test
  void selectorSupportRejectsInvalidScalarFieldsAndA1Geometry() {
    assertEquals(
        "field must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireNonBlank(" ", "field"))
            .getMessage());
    assertEquals(
        "field must not be null",
        assertThrows(
                NullPointerException.class,
                () -> SelectorValueValidation.requireNonBlank(null, "field"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireDefinedName("1bad", "selector"))
            .getMessage()
            .startsWith("selector "));
    assertEquals(
        "name must start with a letter or underscore and contain only letters, digits, underscore, or period",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireDefinedName("1bad", "name"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requirePivotTableName(" ", "selector"))
            .getMessage()
            .startsWith("selector "));
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireAddress("A0", "selector"))
            .getMessage()
            .startsWith("selector "));
    assertEquals(
        "selector must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireRange(" ", "selector"))
            .getMessage());
    assertEquals(
        "selector must be a rectangular A1-style range with at most one ':'",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireRange("A1:B2:C3", "selector"))
            .getMessage());
    assertEquals(
        "count must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requirePositive(0, "count"))
            .getMessage());
    assertEquals(
        "index must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireNonNegative(-1, "index"))
            .getMessage());
    assertEquals(
        "delta must not be 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireNonZero(0, "delta"))
            .getMessage());
    assertEquals(
        "row must be within Excel .xlsx row bounds",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireRowIndexWithinBounds(1_048_576, "row"))
            .getMessage());
    assertEquals(
        "column must be within Excel .xlsx column bounds",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireColumnIndexWithinBounds(16_384, "column"))
            .getMessage());
    assertEquals(
        "rowCount * columnCount must not exceed 250000 but was 250500",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireWindowSize(501, 500))
            .getMessage());
  }

  @Test
  void selectorSupportCopiesDistinctCollectionsAndRejectsDuplicateOrNullEntries() {
    assertEquals(
        List.of("A1", "B2"),
        SelectorListValidation.copyDistinctAddresses(List.of("A1", "B2"), "addresses"));
    assertEquals(
        List.of("A1", "B2"),
        SelectorListValidation.copyDistinctAddresses(List.of("A1", "B2"), "cells"));
    assertEquals(
        List.of("A1:B2"), SelectorListValidation.copyDistinctRanges(List.of("A1:B2"), "ranges"));
    assertEquals(
        List.of("Budget", "Ops"),
        SelectorListValidation.copyDistinctSheetNames(List.of("Budget", "Ops"), "sheetNames"));
    assertEquals(
        List.of("BudgetTotal"),
        SelectorListValidation.copyDistinctDefinedNames(List.of("BudgetTotal"), "names"));
    assertEquals(
        List.of("Sales Pivot 2026"),
        SelectorListValidation.copyDistinctPivotTableNames(List.of("Sales Pivot 2026"), "names"));
    assertEquals(
        List.of("x", "y"), SelectorListValidation.copyDistinctValues(List.of("x", "y"), "values"));
    assertEquals(
        List.of(
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new NamedRangeSelector.SheetScope("LocalItem", "Budget")),
        SelectorListValidation.copyDistinctNamedRangeRefs(
            List.of(
                new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                new NamedRangeSelector.SheetScope("LocalItem", "Budget")),
            "selectors"));

    assertEquals(
        "addresses must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctAddresses(List.of("A1", "A1"), "addresses"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorListValidation.copyDistinctAddresses(List.of("A0"), "cells"))
            .getMessage()
            .startsWith("cells[0] address "));
    assertEquals(
        "ranges must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctRanges(List.of("A1:B2", "A1:B2"), "ranges"))
            .getMessage());
    assertEquals(
        "sheetNames must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctSheetNames(
                        List.of("Budget", "Budget"), "sheetNames"))
            .getMessage());
    assertEquals(
        "names must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctDefinedNames(
                        List.of("BudgetTotal", "BudgetTotal"), "names"))
            .getMessage());
    assertEquals(
        "names must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctPivotTableNames(
                        List.of("Sales Pivot 2026", "sales pivot 2026"), "names"))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorListValidation.copyDistinctPivotTableNames(List.of(" "), "names"))
            .getMessage()
            .startsWith("names[0] "));
    assertEquals(
        "values[1] must not be null",
        assertThrows(
                NullPointerException.class,
                () -> SelectorListValidation.copyDistinctValues(Arrays.asList("x", null), "values"))
            .getMessage());
    assertEquals(
        "values must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorListValidation.copyDistinctValues(List.of(), "values"))
            .getMessage());
    assertEquals(
        "values must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorListValidation.copyDistinctValues(List.of("x", "x"), "values"))
            .getMessage());
    assertEquals(
        "selectors must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorListValidation.copyDistinctNamedRangeRefs(List.of(), "selectors"))
            .getMessage());
    assertEquals(
        "selectors must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctNamedRangeRefs(
                        List.of(
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new NamedRangeSelector.WorkbookScope("budgettotal")),
                        "selectors"))
            .getMessage());
  }

  @Test
  void selectorSupportCoversZeroIterationAddressParsingAndCatalogLookupOrdering() {
    assertEquals(-1, SelectorAddressSupport.columnIndex(""));
    assertEquals(-1, SelectorAddressSupport.columnIndex("123"));
    assertThrows(NumberFormatException.class, () -> SelectorAddressSupport.rowIndex(""));
    assertEquals(122, SelectorAddressSupport.rowIndex("123"));
  }

  @Test
  void prefixedValidationMessagePreservesNullBlankAndAlreadyPrefixedMessages() {
    assertNull(SelectorValueValidation.prefixedValidationMessage("field", null));
    assertEquals(" ", SelectorValueValidation.prefixedValidationMessage("field", " "));
    assertEquals(
        "field must not be blank",
        SelectorValueValidation.prefixedValidationMessage("field", "field must not be blank"));
    assertEquals(
        "field invalid", SelectorValueValidation.prefixedValidationMessage("field", "invalid"));
  }
}
