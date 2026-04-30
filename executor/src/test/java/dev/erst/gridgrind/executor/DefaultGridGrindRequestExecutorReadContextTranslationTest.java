package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Read-type and read-context extraction coverage. */
class DefaultGridGrindRequestExecutorReadContextTranslationTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void readTypeReturnsDiscriminatorsForAllReadVariants() {
    assertEquals(
        List.of(
            "GET_WORKBOOK_SUMMARY",
            "GET_WORKBOOK_PROTECTION",
            "GET_NAMED_RANGES",
            "GET_SHEET_SUMMARY",
            "GET_CELLS",
            "GET_WINDOW",
            "GET_MERGED_REGIONS",
            "GET_HYPERLINKS",
            "GET_COMMENTS",
            "GET_SHEET_LAYOUT",
            "GET_PRINT_LAYOUT",
            "GET_DATA_VALIDATIONS",
            "GET_CONDITIONAL_FORMATTING",
            "GET_AUTOFILTERS",
            "GET_TABLES",
            "GET_FORMULA_SURFACE",
            "GET_SHEET_SCHEMA",
            "GET_NAMED_RANGE_SURFACE",
            "ANALYZE_FORMULA_HEALTH",
            "ANALYZE_DATA_VALIDATION_HEALTH",
            "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
            "ANALYZE_AUTOFILTER_HEALTH",
            "ANALYZE_TABLE_HEALTH",
            "ANALYZE_HYPERLINK_HEALTH",
            "ANALYZE_NAMED_RANGE_HEALTH",
            "ANALYZE_WORKBOOK_FINDINGS"),
        List.of(
            readType(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())),
            readType(
                inspect(
                    "workbook-protection",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookProtection())),
            readType(
                inspect(
                    "ranges",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRanges())),
            readType(
                inspect(
                    "sheet",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetSheetSummary())),
            readType(
                inspect(
                    "cells",
                    new CellSelector.ByAddresses("Budget", List.of("A1")),
                    new InspectionQuery.GetCells())),
            readType(
                inspect(
                    "window",
                    new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
                    new InspectionQuery.GetWindow())),
            readType(
                inspect(
                    "merged",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetMergedRegions())),
            readType(
                inspect(
                    "hyperlinks",
                    new CellSelector.AllUsedInSheet("Budget"),
                    new InspectionQuery.GetHyperlinks())),
            readType(
                inspect(
                    "comments",
                    new CellSelector.AllUsedInSheet("Budget"),
                    new InspectionQuery.GetComments())),
            readType(
                inspect(
                    "layout",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetSheetLayout())),
            readType(
                inspect(
                    "print-layout",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetPrintLayout())),
            readType(
                inspect(
                    "validations",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetDataValidations())),
            readType(
                inspect(
                    "conditional-formatting",
                    new RangeSelector.AllOnSheet("Budget"),
                    new InspectionQuery.GetConditionalFormatting())),
            readType(
                inspect(
                    "autofilters",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetAutofilters())),
            readType(
                inspect(
                    "tables",
                    new TableSelector.ByNames(List.of("BudgetTable")),
                    new InspectionQuery.GetTables())),
            readType(
                inspect(
                    "formula", new SheetSelector.All(), new InspectionQuery.GetFormulaSurface())),
            readType(
                inspect(
                    "schema",
                    new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
                    new InspectionQuery.GetSheetSchema())),
            readType(
                inspect(
                    "surface",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.GetNamedRangeSurface())),
            readType(
                inspect(
                    "formula-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeFormulaHealth())),
            readType(
                inspect(
                    "validation-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeDataValidationHealth())),
            readType(
                inspect(
                    "conditional-formatting-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeConditionalFormattingHealth())),
            readType(
                inspect(
                    "autofilter-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeAutofilterHealth())),
            readType(
                inspect(
                    "table-health",
                    new TableSelector.All(),
                    new InspectionQuery.AnalyzeTableHealth())),
            readType(
                inspect(
                    "hyperlink-health",
                    new SheetSelector.All(),
                    new InspectionQuery.AnalyzeHyperlinkHealth())),
            readType(
                inspect(
                    "named-range-health",
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                    new InspectionQuery.AnalyzeNamedRangeHealth())),
            readType(
                inspect(
                    "workbook-findings",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.AnalyzeWorkbookFindings()))));
  }

  @Test
  void extractsContextForReadOperations() {
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    RuntimeException runtimeException = new RuntimeException("x");

    InspectionStep workbook =
        inspect(
            "workbook", new WorkbookSelector.Current(), new InspectionQuery.GetWorkbookSummary());
    InspectionStep workbookProtection =
        inspect(
            "workbook-protection",
            new WorkbookSelector.Current(),
            new InspectionQuery.GetWorkbookProtection());
    InspectionStep namedRanges =
        inspect(
            "ranges",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
            new InspectionQuery.GetNamedRanges());
    InspectionStep sheet =
        inspect("sheet", new SheetSelector.ByName("Budget"), new InspectionQuery.GetSheetSummary());
    InspectionStep cells =
        inspect(
            "cells",
            new CellSelector.ByAddresses("Budget", List.of("A1")),
            new InspectionQuery.GetCells());
    InspectionStep window =
        inspect(
            "window",
            new RangeSelector.RectangularWindow("Budget", "B2", 2, 2),
            new InspectionQuery.GetWindow());
    InspectionStep merged =
        inspect(
            "merged", new SheetSelector.ByName("Budget"), new InspectionQuery.GetMergedRegions());
    InspectionStep hyperlinks =
        inspect(
            "hyperlinks",
            new CellSelector.AllUsedInSheet("Budget"),
            new InspectionQuery.GetHyperlinks());
    InspectionStep comments =
        inspect(
            "comments",
            new CellSelector.ByAddresses("Budget", List.of("A1")),
            new InspectionQuery.GetComments());
    InspectionStep layout =
        inspect("layout", new SheetSelector.ByName("Budget"), new InspectionQuery.GetSheetLayout());
    InspectionStep printLayout =
        inspect(
            "print-layout",
            new SheetSelector.ByName("Budget"),
            new InspectionQuery.GetPrintLayout());
    InspectionStep validations =
        inspect(
            "validations",
            new RangeSelector.ByRanges("Budget", List.of("A1:A3")),
            new InspectionQuery.GetDataValidations());
    InspectionStep conditionalFormatting =
        inspect(
            "conditional-formatting",
            new RangeSelector.ByRanges("Budget", List.of("B2:B5")),
            new InspectionQuery.GetConditionalFormatting());
    InspectionStep autofilters =
        inspect(
            "autofilters",
            new SheetSelector.ByName("Budget"),
            new InspectionQuery.GetAutofilters());
    InspectionStep tables =
        inspect(
            "tables",
            new TableSelector.ByNames(List.of("BudgetTable")),
            new InspectionQuery.GetTables());
    InspectionStep formula =
        inspect("formula", new SheetSelector.All(), new InspectionQuery.GetFormulaSurface());
    InspectionStep schema =
        inspect(
            "schema",
            new RangeSelector.RectangularWindow("Budget", "C3", 2, 2),
            new InspectionQuery.GetSheetSchema());
    InspectionStep surface =
        inspect(
            "surface",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
            new InspectionQuery.GetNamedRangeSurface());
    InspectionStep formulaHealth =
        inspect(
            "formula-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeFormulaHealth());
    InspectionStep validationHealth =
        inspect(
            "validation-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeDataValidationHealth());
    InspectionStep conditionalFormattingHealth =
        inspect(
            "conditional-formatting-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeConditionalFormattingHealth());
    InspectionStep autofilterHealth =
        inspect(
            "autofilter-health",
            new SheetSelector.ByNames(List.of("Budget")),
            new InspectionQuery.AnalyzeAutofilterHealth());
    InspectionStep tableHealth =
        inspect(
            "table-health",
            new TableSelector.ByNames(List.of("BudgetTable")),
            new InspectionQuery.AnalyzeTableHealth());
    InspectionStep hyperlinkHealth =
        inspect(
            "hyperlink-health",
            new SheetSelector.All(),
            new InspectionQuery.AnalyzeHyperlinkHealth());
    InspectionStep namedRangeHealth =
        inspect(
            "named-range-health",
            new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                List.of(
                    new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                        "BudgetTotal"))),
            new InspectionQuery.AnalyzeNamedRangeHealth());
    InspectionStep workbookFindings =
        inspect(
            "workbook-findings",
            new WorkbookSelector.Current(),
            new InspectionQuery.AnalyzeWorkbookFindings());

    assertReadContext(workbook, null, null, null, runtimeException);
    assertReadContext(workbookProtection, null, null, null, runtimeException);
    assertReadContext(namedRanges, null, null, null, runtimeException);
    assertReadContext(sheet, "Budget", null, null, runtimeException);
    assertReadContext(cells, "Budget", null, null, runtimeException);
    assertReadContext(window, "Budget", "B2", null, runtimeException);
    assertReadContext(merged, "Budget", null, null, runtimeException);
    assertReadContext(hyperlinks, "Budget", null, null, runtimeException);
    assertReadContext(comments, "Budget", null, null, runtimeException);
    assertReadContext(layout, "Budget", null, null, runtimeException);
    assertReadContext(printLayout, "Budget", null, null, runtimeException);
    assertReadContext(validations, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormatting, "Budget", null, null, runtimeException);
    assertReadContext(autofilters, "Budget", null, null, runtimeException);
    assertReadContext(tables, null, null, null, runtimeException);
    assertReadContext(formula, null, null, null, runtimeException);
    assertReadContext(schema, "Budget", "C3", null, runtimeException);
    assertReadContext(surface, null, null, null, runtimeException);
    assertReadContext(formulaHealth, "Budget", null, null, runtimeException);
    assertReadContext(validationHealth, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormattingHealth, "Budget", null, null, runtimeException);
    assertReadContext(autofilterHealth, "Budget", null, null, runtimeException);
    assertReadContext(tableHealth, null, null, null, runtimeException);
    assertReadContext(hyperlinkHealth, null, null, null, runtimeException);
    assertReadContext(namedRangeHealth, null, null, "BudgetTotal", runtimeException);
    assertReadContext(workbookFindings, null, null, null, runtimeException);

    assertEquals("BudgetTotal", namedRangeNameFor(workbook, missingNamedRange));
    assertEquals("BudgetTotal", namedRangeNameFor(namedRanges, missingNamedRange));
    assertEquals("BAD!", addressFor(cells, invalidAddress));
  }

  @Test
  void extractsSingleSheetAndNamedRangeContextOnlyWhenSelectionsAreUnambiguous() {
    assertEquals(
        "Budget",
        sheetNameFor(
            inspect(
                "formula",
                new SheetSelector.ByNames(List.of("Budget")),
                new InspectionQuery.GetFormulaSurface())));
    assertNull(
        sheetNameFor(
            inspect(
                "hyperlink-health",
                new SheetSelector.ByNames(List.of("Budget", "Forecast")),
                new InspectionQuery.AnalyzeHyperlinkHealth())));

    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName(
                            "BudgetTotal"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope(
                            "BudgetTotal", "Budget"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
    assertNull(
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf(
                    List.of(
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                            "BudgetTotal"),
                        new dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope(
                            "ForecastTotal"))),
                new InspectionQuery.GetNamedRangeSurface()),
            new RuntimeException("x")));
  }

  @Test
  void extractsReadContextFromExceptionsBeforeFallingBackToReadShape() {
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());

    assertEquals(
        "C3",
        addressFor(
            inspect(
                "schema",
                new RangeSelector.RectangularWindow("Budget", "C3", 2, 2),
                new InspectionQuery.GetSheetSchema()),
            invalidAddress));
    assertEquals(
        "BudgetTotal",
        namedRangeNameFor(
            inspect(
                "surface",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.GetNamedRangeSurface()),
            missingNamedRange));
  }
}
