package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.problem.FieldValidationAddressRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.GridGrindRequestProblemSupport;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelReadLimits;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
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
    assertEquals(
        "Budget_Total", SelectorValueValidation.requireDefinedName("Budget_Total", "selector"));
    assertEquals(
        "Sales Pivot 2026",
        SelectorValueValidation.requirePivotTableName("Sales Pivot 2026", "selector"));
    assertEquals(
        "A1:B3",
        SelectorValueValidation.requireOrderedRange(new RangeRef("A1:B3"), "selector").value());
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
        "name must start with a letter, underscore, or backslash and contain only Unicode letters, Unicode numbers, underscore, period, or backslash",
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
        "row must not exceed 1048575 (Excel row limit)",
        assertThrows(
                IllegalArgumentException.class,
                () -> SelectorValueValidation.requireRowIndexWithinBounds(1_048_576, "row"))
            .getMessage());
    assertEquals(
        "column must not exceed 16383 (Excel column limit)",
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
    assertEquals(
        "addresses must not exceed 250000 but was 250001",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctAddresses(
                        exactAddresses(ExcelReadLimits.MAX_READ_CELLS + 1).stream()
                            .map(CellAddress::new)
                            .toList(),
                        "addresses"))
            .getMessage());
  }

  @Test
  void selectorSupportCopiesDistinctCollectionsAndRejectsDuplicateOrNullEntries() {
    assertEquals(
        List.of("A1", "B2"),
        SelectorListValidation.copyDistinctAddresses(
                List.of(new CellAddress("A1"), new CellAddress("B2")), "addresses")
            .stream()
            .map(CellAddress::value)
            .toList());
    assertEquals(
        List.of("A1", "B2"),
        SelectorListValidation.copyDistinctAddresses(
                List.of(new CellAddress("A1"), new CellAddress("B2")), "cells")
            .stream()
            .map(CellAddress::value)
            .toList());
    assertEquals(
        List.of("A1:B2"),
        SelectorListValidation.copyDistinctRanges(List.of(new RangeRef("A1:B2")), "ranges").stream()
            .map(RangeRef::value)
            .toList());
    assertEquals(
        List.of("Budget", "Ops"),
        SelectorListValidation.copyDistinctSheetNames(
                List.of(new SheetName("Budget"), new SheetName("Ops")), "sheetNames")
            .stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
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
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
            new NamedRangeSelector.SheetScope(
                new DefinedName("LocalItem"), new SheetName("Budget"))),
        SelectorListValidation.copyDistinctNamedRangeRefs(
            List.of(
                new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
                new NamedRangeSelector.SheetScope(
                    new DefinedName("LocalItem"), new SheetName("Budget"))),
            "selectors"));

    assertEquals(
        "addresses must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctAddresses(
                        List.of(new CellAddress("A1"), new CellAddress("A1")), "addresses"))
            .getMessage());
    assertEquals(
        "ranges must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctRanges(
                        List.of(new RangeRef("A1:B2"), new RangeRef("A1:B2")), "ranges"))
            .getMessage());
    assertEquals(
        "sheetNames must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorListValidation.copyDistinctSheetNames(
                        List.of(new SheetName("Budget"), new SheetName("Budget")), "sheetNames"))
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
                            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
                            new NamedRangeSelector.WorkbookScope(new DefinedName("budgettotal"))),
                        "selectors"))
            .getMessage());
  }

  private static List<String> exactAddresses(int count) {
    return java.util.stream.IntStream.rangeClosed(1, count).mapToObj(index -> "A" + index).toList();
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

  @Test
  void cellAddressCarriesTypedSyntaxGuidance() {
    var failure =
        assertThrows(
            dev.erst.gridgrind.core.problem.InvalidRequestException.class,
            () -> new CellAddress("A0"));
    FieldValidationProblem requestProblem =
        assertInstanceOf(FieldValidationProblem.class, failure.requestProblem());

    assertEquals("address must be a single-cell A1-style address", failure.getMessage());
    assertEquals(FieldValidationAddressRule.ADDRESS_SYNTAX, requestProblem.rule());
    assertEquals(
        "Use a single-cell A1-style address such as A1 or BC12 within Excel .xlsx bounds for field 'address'.",
        GridGrindRequestProblemSupport.resolution(
            requestProblem,
            new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.standardInput(),
                ProblemContextRequestSurfaces.JsonLocation.unavailable())));
  }
}
