package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports;
import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports.AnalysisFindingReport;
import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports.AnalysisLocationReport;
import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports.AnalysisSummaryReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage lock-in for Phase-4 assertion execution, helper seams, and composite failures. */
class AssertionExecutorCoverageTest {
  @Test
  void executesLeafAssertionsAcrossWorkbookFactsAndAnalysis() throws IOException {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Path workbookPath = Files.createTempFile("gridgrind-assertions-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    success(
        executor.execute(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                List.of(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(textCell("Owner"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B1"),
                        new CellMutationAction.SetCell(textCell("Amount"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A2"),
                        new CellMutationAction.SetCell(textCell("Ada"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B2"),
                        new CellMutationAction.SetCell(new CellInput.Numeric(42.0d))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "C2"),
                        new CellMutationAction.SetCell(new CellInput.BooleanValue(true))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "E2"),
                        new CellMutationAction.SetCell(formulaCell("1/0"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "F2"),
                        new CellMutationAction.SetCell(formulaCell("2+3"))),
                    mutate(
                        new WorkbookSelector.Current(),
                        new WorkbookMutationAction.SetWorkbookProtection(
                            new WorkbookProtectionInput(true, false, false, null, null))),
                    mutate(
                        new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                        new StructuredMutationAction.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            new NamedRangeTarget("Budget", "F2"))),
                    mutate(
                        new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                        new StructuredMutationAction.SetTable(
                            TableInput.withDefaultMetadata(
                                "BudgetTable",
                                "Budget",
                                "A1:B2",
                                false,
                                new TableStyleInput.None())))),
                List.of(),
                List.of())));

    GridGrindResponse.Success inspected =
        success(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(calculateAll()),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses(
                                "Budget", List.of("A2", "B2", "C2", "D2", "E2", "F2")),
                            new InspectionQuery.GetCells()),
                        inspect(
                            "protection",
                            new WorkbookSelector.Current(),
                            new InspectionQuery.GetWorkbookProtection()),
                        inspect(
                            "sheet",
                            new SheetSelector.ByName("Budget"),
                            new InspectionQuery.GetSheetSummary()),
                        inspect(
                            "namedRanges",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new InspectionQuery.GetNamedRanges()),
                        inspect(
                            "tables",
                            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                            new InspectionQuery.GetTables()),
                        inspect(
                            "formulaHealth",
                            new SheetSelector.ByName("Budget"),
                            new InspectionQuery.AnalyzeFormulaHealth())))));

    InspectionResult.CellsResult cells =
        inspection(inspected, "cells", InspectionResult.CellsResult.class);
    InspectionResult.WorkbookProtectionResult protection =
        inspection(inspected, "protection", InspectionResult.WorkbookProtectionResult.class);
    InspectionResult.SheetSummaryResult sheet =
        inspection(inspected, "sheet", InspectionResult.SheetSummaryResult.class);
    InspectionResult.NamedRangesResult namedRanges =
        inspection(inspected, "namedRanges", InspectionResult.NamedRangesResult.class);
    InspectionResult.TablesResult tables =
        inspection(inspected, "tables", InspectionResult.TablesResult.class);
    InspectionResult.FormulaHealthResult formulaHealth =
        inspection(inspected, "formulaHealth", InspectionResult.FormulaHealthResult.class);

    dev.erst.gridgrind.contract.dto.CellReport.TextReport owner =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().get(0));
    dev.erst.gridgrind.contract.dto.CellReport.NumberReport amount =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, cells.cells().get(1));
    dev.erst.gridgrind.contract.dto.CellReport.BooleanReport enabled =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.BooleanReport.class, cells.cells().get(2));
    assertInstanceOf(
        dev.erst.gridgrind.contract.dto.CellReport.BlankReport.class, cells.cells().get(3));
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport errorFormula =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class, cells.cells().get(4));
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport totalFormula =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class, cells.cells().get(5));
    AnalysisFindingReport firstFinding = formulaHealth.analysis().findings().getFirst();

    GridGrindResponse.Success asserted =
        success(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "present-named-range",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new Assertion.NamedRangePresent()),
                        assertThat(
                            "absent-named-range",
                            new NamedRangeSelector.WorkbookScope("MissingTotal"),
                            new Assertion.NamedRangeAbsent()),
                        assertThat(
                            "present-table",
                            new TableSelector.ByName("BudgetTable"),
                            new Assertion.TablePresent()),
                        assertThat(
                            "cell-text",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Text(
                                    owner.stringValue()))),
                        assertThat(
                            "cell-number",
                            new CellSelector.ByAddress("Budget", "B2"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                    .NumericValue(amount.numberValue()))),
                        assertThat(
                            "cell-boolean",
                            new CellSelector.ByAddress("Budget", "C2"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                    .BooleanValue(enabled.booleanValue()))),
                        assertThat(
                            "cell-blank",
                            new CellSelector.ByAddress("Budget", "D2"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                    .Blank())),
                        assertThat(
                            "cell-error",
                            new CellSelector.ByAddress("Budget", "E2"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                    .ErrorValue(
                                    assertInstanceOf(
                                            dev.erst.gridgrind.contract.dto.CellReport.ErrorReport
                                                .class,
                                            errorFormula.evaluation())
                                        .errorValue()))),
                        assertThat(
                            "cell-formula-evaluation",
                            new CellSelector.ByAddress("Budget", "F2"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                    .NumericValue(
                                    assertInstanceOf(
                                            dev.erst.gridgrind.contract.dto.CellReport.NumberReport
                                                .class,
                                            totalFormula.evaluation())
                                        .numberValue()))),
                        assertThat(
                            "display-value",
                            new CellSelector.ByAddress("Budget", "B2"),
                            new Assertion.DisplayValue(amount.displayValue())),
                        assertThat(
                            "formula-text",
                            new CellSelector.ByAddress("Budget", "F2"),
                            new Assertion.FormulaText(totalFormula.formula())),
                        assertThat(
                            "cell-style",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new Assertion.CellStyle(owner.style())),
                        assertThat(
                            "workbook-protection",
                            new WorkbookSelector.Current(),
                            new Assertion.WorkbookProtectionFacts(protection.protection())),
                        assertThat(
                            "sheet-structure",
                            new SheetSelector.ByName("Budget"),
                            new Assertion.SheetStructureFacts(sheet.sheet())),
                        assertThat(
                            "named-range-facts",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new Assertion.NamedRangeFacts(namedRanges.namedRanges())),
                        assertThat(
                            "table-facts",
                            new TableSelector.ByName("BudgetTable"),
                            new Assertion.TableFacts(tables.tables())),
                        assertThat(
                            "analysis-max-severity",
                            new SheetSelector.ByName("Budget"),
                            new Assertion.AnalysisMaxSeverity(
                                new InspectionQuery.AnalyzeFormulaHealth(),
                                highestSeverity(formulaHealth.analysis().summary()))),
                        assertThat(
                            "analysis-finding-present",
                            new SheetSelector.ByName("Budget"),
                            new Assertion.AnalysisFindingPresent(
                                new InspectionQuery.AnalyzeFormulaHealth(),
                                firstFinding.code(),
                                firstFinding.severity(),
                                firstFinding.message().substring(0, 3))),
                        assertThat(
                            "analysis-finding-absent",
                            new SheetSelector.ByName("Budget"),
                            new Assertion.AnalysisFindingAbsent(
                                new InspectionQuery.AnalyzeFormulaHealth(),
                                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                                null,
                                null)),
                        assertThat(
                            "all-of",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new Assertion.AllOf(
                                List.of(
                                    new Assertion.CellValue(
                                        new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                            .Text(owner.stringValue())),
                                    new Assertion.DisplayValue(owner.displayValue())))),
                        assertThat(
                            "any-of",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new Assertion.AnyOf(
                                List.of(
                                    new Assertion.DisplayValue("Wrong"),
                                    new Assertion.CellValue(
                                        new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                            .Text(owner.stringValue()))))),
                        assertThat(
                            "not",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new Assertion.Not(
                                new Assertion.CellValue(
                                    new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                        .Text("Wrong"))))),
                    List.of())));

    assertFalse(asserted.assertions().isEmpty());
    assertEquals(
        List.of(
            "present-named-range",
            "absent-named-range",
            "present-table",
            "cell-text",
            "cell-number",
            "cell-boolean",
            "cell-blank",
            "cell-error",
            "cell-formula-evaluation",
            "display-value",
            "formula-text",
            "cell-style",
            "workbook-protection",
            "sheet-structure",
            "named-range-facts",
            "table-facts",
            "analysis-max-severity",
            "analysis-finding-present",
            "analysis-finding-absent",
            "all-of",
            "any-of",
            "not"),
        asserted.assertions().stream()
            .map(dev.erst.gridgrind.contract.assertion.AssertionResult::stepId)
            .toList());
  }

  @Test
  void factAndAnalysisAssertionFailuresReturnStructuredProblems() throws IOException {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Path workbookPath = Files.createTempFile("gridgrind-assertion-family-failures-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    success(
        executor.execute(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                List.of(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(textCell("Owner"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B1"),
                        new CellMutationAction.SetCell(textCell("Amount"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A2"),
                        new CellMutationAction.SetCell(textCell("Ada"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "B2"),
                        new CellMutationAction.SetCell(new CellInput.Numeric(42.0d))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "E2"),
                        new CellMutationAction.SetCell(formulaCell("1/0"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "F2"),
                        new CellMutationAction.SetCell(formulaCell("2+3"))),
                    mutate(
                        new WorkbookSelector.Current(),
                        new WorkbookMutationAction.SetWorkbookProtection(
                            new WorkbookProtectionInput(true, false, false, null, null))),
                    mutate(
                        new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                        new StructuredMutationAction.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            new NamedRangeTarget("Budget", "F2"))),
                    mutate(
                        new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                        new StructuredMutationAction.SetTable(
                            TableInput.withDefaultMetadata(
                                "BudgetTable",
                                "Budget",
                                "A1:B2",
                                false,
                                new TableStyleInput.None())))),
                List.of(),
                List.of())));

    GridGrindResponse.Success inspected =
        success(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(calculateAll()),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "cells",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new InspectionQuery.GetCells()),
                        inspect(
                            "protection",
                            new WorkbookSelector.Current(),
                            new InspectionQuery.GetWorkbookProtection()),
                        inspect(
                            "sheet",
                            new SheetSelector.ByName("Budget"),
                            new InspectionQuery.GetSheetSummary()),
                        inspect(
                            "namedRanges",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new InspectionQuery.GetNamedRanges()),
                        inspect(
                            "tables",
                            new TableSelector.ByName("BudgetTable"),
                            new InspectionQuery.GetTables()),
                        inspect(
                            "formulaHealth",
                            new SheetSelector.ByName("Budget"),
                            new InspectionQuery.AnalyzeFormulaHealth())))));

    dev.erst.gridgrind.contract.dto.CellReport.TextReport owner =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
            inspection(inspected, "cells", InspectionResult.CellsResult.class).cells().getFirst());
    InspectionResult.WorkbookProtectionResult protection =
        inspection(inspected, "protection", InspectionResult.WorkbookProtectionResult.class);
    InspectionResult.SheetSummaryResult sheet =
        inspection(inspected, "sheet", InspectionResult.SheetSummaryResult.class);
    InspectionResult.NamedRangesResult namedRanges =
        inspection(inspected, "namedRanges", InspectionResult.NamedRangesResult.class);
    InspectionResult.TablesResult tables =
        inspection(inspected, "tables", InspectionResult.TablesResult.class);
    InspectionResult.FormulaHealthResult formulaHealth =
        inspection(inspected, "formulaHealth", InspectionResult.FormulaHealthResult.class);
    AnalysisFindingReport firstFinding = formulaHealth.analysis().findings().getFirst();

    GridGrindResponse.Failure presentMissing =
        assertionFailure(
            executor,
            workbookPath,
            "present-missing-range",
            new NamedRangeSelector.WorkbookScope("MissingTotal"),
            new Assertion.NamedRangePresent());
    assertTrue(presentMissing.problem().message().contains("EXPECT_NAMED_RANGE_PRESENT"));
    assertEquals(
        List.of(),
        assertInstanceOf(
                InspectionResult.NamedRangesResult.class,
                presentMissing.problem().assertionFailure().orElseThrow().observations().getFirst())
            .namedRanges());

    GridGrindResponse.Failure absentTable =
        assertionFailure(
            executor,
            workbookPath,
            "absent-table",
            new TableSelector.ByName("BudgetTable"),
            new Assertion.TableAbsent());
    assertTrue(absentTable.problem().message().contains("EXPECT_TABLE_ABSENT"));

    GridGrindResponse.Failure styleMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "style-mismatch",
            new CellSelector.ByAddress("Budget", "A2"),
            new Assertion.CellStyle(
                new GridGrindWorkbookSurfaceReports.CellStyleReport(
                    "0",
                    owner.style().alignment(),
                    owner.style().font(),
                    owner.style().fill(),
                    owner.style().border(),
                    owner.style().protection())));
    assertTrue(styleMismatch.problem().message().contains("EXPECT_CELL_STYLE"));

    GridGrindResponse.Failure formulaTextMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "formula-text-mismatch",
            new CellSelector.ByAddress("Budget", "F2"),
            new Assertion.FormulaText("1+1"));
    assertTrue(formulaTextMismatch.problem().message().contains("EXPECT_FORMULA_TEXT"));

    WorkbookProtectionReport expectedProtection = protection.protection();
    GridGrindResponse.Failure protectionMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "workbook-protection-mismatch",
            new WorkbookSelector.Current(),
            new Assertion.WorkbookProtectionFacts(
                new WorkbookProtectionReport(
                    !expectedProtection.structureLocked(),
                    expectedProtection.windowsLocked(),
                    expectedProtection.revisionsLocked(),
                    expectedProtection.workbookPasswordHashPresent(),
                    expectedProtection.revisionsPasswordHashPresent())));
    assertTrue(protectionMismatch.problem().message().contains("EXPECT_WORKBOOK_PROTECTION"));

    GridGrindWorkbookSurfaceReports.SheetSummaryReport expectedSheet = sheet.sheet();
    GridGrindResponse.Failure sheetMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "sheet-structure-mismatch",
            new SheetSelector.ByName("Budget"),
            new Assertion.SheetStructureFacts(
                new GridGrindWorkbookSurfaceReports.SheetSummaryReport(
                    expectedSheet.sheetName(),
                    expectedSheet.visibility(),
                    expectedSheet.protection(),
                    expectedSheet.physicalRowCount(),
                    expectedSheet.lastRowIndex() + 1,
                    expectedSheet.lastColumnIndex())));
    assertTrue(sheetMismatch.problem().message().contains("EXPECT_SHEET_STRUCTURE"));

    GridGrindResponse.Failure namedRangeMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "named-range-facts-mismatch",
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new Assertion.NamedRangeFacts(List.of()));
    assertTrue(namedRangeMismatch.problem().message().contains("EXPECT_NAMED_RANGE_FACTS"));
    assertEquals(
        namedRanges.namedRanges(),
        assertInstanceOf(
                InspectionResult.NamedRangesResult.class,
                namedRangeMismatch
                    .problem()
                    .assertionFailure()
                    .orElseThrow()
                    .observations()
                    .getFirst())
            .namedRanges());

    GridGrindResponse.Failure tableMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "table-facts-mismatch",
            new TableSelector.ByName("BudgetTable"),
            new Assertion.TableFacts(List.of()));
    assertTrue(tableMismatch.problem().message().contains("EXPECT_TABLE_FACTS"));
    assertEquals(
        tables.tables(),
        assertInstanceOf(
                InspectionResult.TablesResult.class,
                tableMismatch.problem().assertionFailure().orElseThrow().observations().getFirst())
            .tables());

    GridGrindResponse.Failure severityMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "analysis-max-severity-mismatch",
            new SheetSelector.ByName("Budget"),
            new Assertion.AnalysisMaxSeverity(
                new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING));
    assertTrue(severityMismatch.problem().message().contains("EXPECT_ANALYSIS_MAX_SEVERITY"));

    GridGrindResponse.Failure missingFinding =
        assertionFailure(
            executor,
            workbookPath,
            "analysis-finding-present-missing",
            new SheetSelector.ByName("Budget"),
            new Assertion.AnalysisFindingPresent(
                new InspectionQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                null,
                null));
    assertTrue(missingFinding.problem().message().contains("missing finding"));

    GridGrindResponse.Failure unexpectedFinding =
        assertionFailure(
            executor,
            workbookPath,
            "analysis-finding-absent-unexpected",
            new SheetSelector.ByName("Budget"),
            new Assertion.AnalysisFindingAbsent(
                new InspectionQuery.AnalyzeFormulaHealth(),
                firstFinding.code(),
                firstFinding.severity(),
                firstFinding.message().substring(0, 3)));
    assertTrue(unexpectedFinding.problem().message().contains("unexpectedly present"));
  }

  @Test
  void compositeFailuresAndFormulaTextMismatchesReturnStructuredProblems() throws IOException {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Path workbookPath = Files.createTempFile("gridgrind-assertion-failures-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    success(
        executor.execute(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                List.of(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(textCell("Owner")))),
                List.of(),
                List.of())));

    GridGrindResponse.Failure formulaMismatch =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "formula-mismatch",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new Assertion.FormulaText("1+1"))),
                    List.of())));
    assertTrue(formulaMismatch.problem().message().contains("EXPECT_FORMULA_TEXT"));

    GridGrindResponse.Failure allOfFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "all-of-failure",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new Assertion.AllOf(
                                List.of(
                                    new Assertion.CellValue(
                                        new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                            .Text("Owner")),
                                    new Assertion.DisplayValue("Wrong"))))),
                    List.of())));
    assertTrue(allOfFailure.problem().message().contains("ALL_OF failed"));

    GridGrindResponse.Failure anyOfFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "any-of-failure",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new Assertion.AnyOf(
                                List.of(
                                    new Assertion.DisplayValue("Wrong"),
                                    new Assertion.CellValue(
                                        new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                            .Text("Also wrong")))))),
                    List.of())));
    assertTrue(anyOfFailure.problem().message().contains("ANY_OF failed"));

    GridGrindResponse.Failure notFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "not-failure",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new Assertion.Not(
                                new Assertion.CellValue(
                                    new dev.erst.gridgrind.contract.assertion.ExpectedCellValue
                                        .Text("Owner"))))),
                    List.of())));
    assertTrue(notFailure.problem().message().contains("NOT failed"));
    assertTrue(notFailure.problem().assertionFailure().isPresent());
  }

  @Test
  void zeroMatchDisplayFormulaAndStyleAssertionsReturnStructuredProblems() {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    GridGrindResponse.Failure displayFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    List.of(
                        assertThat(
                            "display-missing-table-cell",
                            missingAmountCellTarget(),
                            new Assertion.DisplayValue("999"))),
                    List.of())));
    assertTrue(
        displayFailure
            .problem()
            .message()
            .contains("EXPECT_DISPLAY_VALUE resolved no matching cells"));
    assertEquals(
        "display-missing-table-cell",
        displayFailure.problem().assertionFailure().orElseThrow().stepId());
    assertEquals(
        List.of(),
        assertInstanceOf(
                InspectionResult.CellsResult.class,
                displayFailure.problem().assertionFailure().orElseThrow().observations().getFirst())
            .cells());

    GridGrindResponse.Failure formulaFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    List.of(
                        assertThat(
                            "formula-missing-table-cell",
                            missingAmountCellTarget(),
                            new Assertion.FormulaText("SUM(A1:A2)"))),
                    List.of())));
    assertTrue(
        formulaFailure
            .problem()
            .message()
            .contains("EXPECT_FORMULA_TEXT resolved no matching cells"));
    assertEquals(
        "formula-missing-table-cell",
        formulaFailure.problem().assertionFailure().orElseThrow().stepId());

    GridGrindResponse.Failure styleFailure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    List.of(
                        assertThat(
                            "style-missing-table-cell",
                            missingAmountCellTarget(),
                            new Assertion.CellStyle(style()))),
                    List.of())));
    assertTrue(
        styleFailure.problem().message().contains("EXPECT_CELL_STYLE resolved no matching cells"));
    assertEquals(
        "style-missing-table-cell",
        styleFailure.problem().assertionFailure().orElseThrow().stepId());
  }

  @Test
  void executesChartAndPivotFactAssertionsFromPublicWorkflows() throws IOException {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Path chartPath = Files.createTempFile("gridgrind-chart-assertions-", ".xlsx");
    Path pivotPath = Files.createTempFile("gridgrind-pivot-assertions-", ".xlsx");
    Files.deleteIfExists(chartPath);
    Files.deleteIfExists(pivotPath);

    success(executor.execute(rewritePersistence(readExample("chart-request.json"), chartPath)));
    success(executor.execute(rewritePersistence(readExample("pivot-request.json"), pivotPath)));

    InspectionResult.ChartsResult charts =
        inspection(
            success(
                executor.execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(chartPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            inspect(
                                "charts",
                                new ChartSelector.AllOnSheet("Ops"),
                                new InspectionQuery.GetCharts()))))),
            "charts",
            InspectionResult.ChartsResult.class);
    InspectionResult.PivotTablesResult pivots =
        inspection(
            success(
                executor.execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(pivotPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            inspect(
                                "pivots",
                                new PivotTableSelector.All(),
                                new InspectionQuery.GetPivotTables()))))),
            "pivots",
            InspectionResult.PivotTablesResult.class);

    GridGrindResponse.Success asserted =
        success(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(chartPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "chart-facts",
                            new ChartSelector.AllOnSheet("Ops"),
                            new Assertion.ChartFacts(charts.charts())),
                        assertThat(
                            "chart-present",
                            new ChartSelector.ByName("Ops", charts.charts().getFirst().name()),
                            new Assertion.ChartPresent()),
                        assertThat(
                            "chart-absent",
                            new ChartSelector.ByName("Ops", "MissingChart"),
                            new Assertion.ChartAbsent())),
                    List.of())));
    assertEquals(3, asserted.assertions().size());

    GridGrindResponse.Success pivotAssertions =
        success(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(pivotPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "pivot-facts",
                            new PivotTableSelector.All(),
                            new Assertion.PivotTableFacts(pivots.pivotTables())),
                        assertThat(
                            "pivot-present",
                            new PivotTableSelector.ByName(pivots.pivotTables().getFirst().name()),
                            new Assertion.PivotTablePresent()),
                        assertThat(
                            "pivot-absent",
                            new PivotTableSelector.ByName("Missing Pivot"),
                            new Assertion.PivotTableAbsent())),
                    List.of())));
    assertEquals(3, pivotAssertions.assertions().size());

    GridGrindResponse.Failure chartFactsMismatch =
        assertionFailure(
            executor,
            chartPath,
            "chart-facts-mismatch",
            new ChartSelector.AllOnSheet("Ops"),
            new Assertion.ChartFacts(List.of()));
    assertTrue(chartFactsMismatch.problem().message().contains("EXPECT_CHART_FACTS"));

    GridGrindResponse.Failure pivotFactsMismatch =
        assertionFailure(
            executor,
            pivotPath,
            "pivot-facts-mismatch",
            new PivotTableSelector.All(),
            new Assertion.PivotTableFacts(List.of()));
    assertTrue(pivotFactsMismatch.problem().message().contains("EXPECT_PIVOT_TABLE_FACTS"));
  }

  @Test
  void reflectionCoverageExhaustsPrivateHelperBranches() throws Exception {
    var readExecutor = new dev.erst.gridgrind.excel.WorkbookReadExecutor();
    AssertionExecutor assertionExecutor =
        new AssertionExecutor(readExecutor, new SemanticSelectorResolver(readExecutor));

    IllegalArgumentException unsupportedPresence =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                assertionExecutor.presenceObservation(
                    "presence", new WorkbookSelector.Current(), null, null));
    assertTrue(unsupportedPresence.getMessage().contains("Unsupported presence assertion target"));

    IllegalArgumentException unsupportedCharts =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                assertionExecutor.chartsObservation(
                    "charts", new WorkbookSelector.Current(), null, null));
    assertEquals("Unsupported chart assertion target", unsupportedCharts.getMessage());

    IllegalArgumentException unsupportedObservedCount =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                AssertionExecutor.observedCount(
                    new InspectionResult.WorkbookSummaryResult(
                        "summary",
                        new GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets(
                            1, List.of("Budget"), "Budget", List.of("Budget"), 0, false))));
    assertTrue(
        unsupportedObservedCount.getMessage().contains("Unsupported presence observation result"));
    assertEquals(
        List.of(),
        assertInstanceOf(
                InspectionResult.NamedRangesResult.class,
                AssertionExecutor.zeroMatchPresenceObservation(
                    "missing-range", new NamedRangeSelector.WorkbookScope("MissingTotal")))
            .namedRanges());
    assertEquals(
        List.of(),
        assertInstanceOf(
                InspectionResult.TablesResult.class,
                AssertionExecutor.zeroMatchPresenceObservation(
                    "missing-table", new TableSelector.ByName("MissingTable")))
            .tables());
    assertEquals(
        List.of(),
        assertInstanceOf(
                InspectionResult.PivotTablesResult.class,
                AssertionExecutor.zeroMatchPresenceObservation(
                    "missing-pivot", new PivotTableSelector.ByName("Missing Pivot")))
            .pivotTables());
    assertEquals(
        "MissingSheet",
        assertInstanceOf(
                InspectionResult.ChartsResult.class,
                AssertionExecutor.zeroMatchPresenceObservation(
                    "missing-chart-sheet", new ChartSelector.AllOnSheet("MissingSheet")))
            .sheetName());
    assertEquals(
        "MissingSheet",
        assertInstanceOf(
                InspectionResult.ChartsResult.class,
                AssertionExecutor.zeroMatchPresenceObservation(
                    "missing-chart-name", new ChartSelector.ByName("MissingSheet", "MissingChart")))
            .sheetName());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            AssertionExecutor.zeroMatchPresenceObservation(
                "unsupported-zero-match", new WorkbookSelector.Current()));

    AnalysisFindingReport finding =
        new AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_ERROR_RESULT,
            AnalysisSeverity.ERROR,
            "Formula error",
            "Division by zero",
            new AnalysisLocationReport.Cell("Budget", "E2"),
            List.of("E2"));
    AnalysisSummaryReport summary = new AnalysisSummaryReport(1, 1, 0, 0);

    List<InspectionResult.Analysis> analyses =
        List.of(
            new InspectionResult.FormulaHealthResult(
                "formula",
                new GridGrindAnalysisReports.FormulaHealthReport(1, summary, List.of(finding))),
            new InspectionResult.DataValidationHealthResult(
                "validation",
                new dev.erst.gridgrind.contract.dto.DataValidationHealthReport(
                    1, summary, List.of(finding))),
            new InspectionResult.ConditionalFormattingHealthResult(
                "formatting",
                new dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport(
                    1, summary, List.of(finding))),
            new InspectionResult.AutofilterHealthResult(
                "autofilter",
                new dev.erst.gridgrind.contract.dto.AutofilterHealthReport(
                    1, summary, List.of(finding))),
            new InspectionResult.TableHealthResult(
                "table",
                new dev.erst.gridgrind.contract.dto.TableHealthReport(
                    1, summary, List.of(finding))),
            new InspectionResult.PivotTableHealthResult(
                "pivot",
                new dev.erst.gridgrind.contract.dto.PivotTableHealthReport(
                    1, summary, List.of(finding))),
            new InspectionResult.HyperlinkHealthResult(
                "hyperlink",
                new GridGrindAnalysisReports.HyperlinkHealthReport(1, summary, List.of(finding))),
            new InspectionResult.NamedRangeHealthResult(
                "namedRange",
                new GridGrindAnalysisReports.NamedRangeHealthReport(1, summary, List.of(finding))),
            new InspectionResult.WorkbookFindingsResult(
                "workbook",
                new GridGrindAnalysisReports.WorkbookFindingsReport(summary, List.of(finding))));

    for (InspectionResult.Analysis analysis : analyses) {
      assertEquals(summary, AssertionExecutor.analysisSummary(analysis));
      assertEquals(List.of(finding), AssertionExecutor.analysisFindings(analysis));
      assertEquals(
          java.util.Optional.of(AnalysisSeverity.ERROR),
          AssertionExecutor.highestSeverity(analysis));
    }

    assertEquals(-1, AssertionExecutor.severityRank(java.util.Optional.empty()));
    assertEquals(0, AssertionExecutor.severityRank(java.util.Optional.of(AnalysisSeverity.INFO)));
    assertEquals(
        1, AssertionExecutor.severityRank(java.util.Optional.of(AnalysisSeverity.WARNING)));
    assertEquals(2, AssertionExecutor.severityRank(java.util.Optional.of(AnalysisSeverity.ERROR)));
    assertEquals(
        java.util.Optional.of(AnalysisSeverity.WARNING),
        AssertionExecutor.highestSeverity(
            new InspectionResult.FormulaHealthResult(
                "warning",
                new GridGrindAnalysisReports.FormulaHealthReport(
                    1, new AnalysisSummaryReport(1, 0, 1, 0), List.of()))));
    assertEquals(
        java.util.Optional.of(AnalysisSeverity.INFO),
        AssertionExecutor.highestSeverity(
            new InspectionResult.FormulaHealthResult(
                "info",
                new GridGrindAnalysisReports.FormulaHealthReport(
                    1, new AnalysisSummaryReport(1, 0, 0, 1), List.of()))));
    assertEquals(
        java.util.Optional.empty(),
        AssertionExecutor.highestSeverity(
            new InspectionResult.FormulaHealthResult(
                "clean",
                new GridGrindAnalysisReports.FormulaHealthReport(
                    1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    GridGrindWorkbookSurfaceReports.CellStyleReport style = style();
    dev.erst.gridgrind.contract.dto.CellReport.BlankReport blankCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BlankReport(
            "A1", "BLANK", "", style, java.util.Optional.empty(), java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.CellReport.TextReport textCell =
        new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
            "A2",
            "STRING",
            "Owner",
            style,
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            "Owner",
            java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.CellReport.NumberReport numberCell =
        new dev.erst.gridgrind.contract.dto.CellReport.NumberReport(
            "B2",
            "NUMERIC",
            "42",
            style,
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            42.0d);
    dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BooleanReport(
            "C2",
            "BOOLEAN",
            "TRUE",
            style,
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            true);
    dev.erst.gridgrind.contract.dto.CellReport.ErrorReport errorCell =
        new dev.erst.gridgrind.contract.dto.CellReport.ErrorReport(
            "D2",
            "ERROR",
            "#DIV/0!",
            style,
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            "#DIV/0!");
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaCell =
        new dev.erst.gridgrind.contract.dto.CellReport.FormulaReport(
            "E2",
            "FORMULA",
            "42",
            style,
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            "2+40",
            numberCell);

    assertTrue(AssertionExecutor.matchesCellValue(blankCell, new ExpectedCellValue.Blank()));
    assertFalse(AssertionExecutor.matchesCellValue(textCell, new ExpectedCellValue.Blank()));
    assertFalse(AssertionExecutor.matchesCellValue(blankCell, new ExpectedCellValue.Text("Owner")));
    assertTrue(AssertionExecutor.matchesCellValue(textCell, new ExpectedCellValue.Text("Owner")));
    assertFalse(AssertionExecutor.matchesCellValue(textCell, new ExpectedCellValue.Text("Wrong")));
    assertFalse(
        AssertionExecutor.matchesCellValue(textCell, new ExpectedCellValue.NumericValue(42.0d)));
    assertTrue(
        AssertionExecutor.matchesCellValue(numberCell, new ExpectedCellValue.NumericValue(42.0d)));
    assertFalse(
        AssertionExecutor.matchesCellValue(numberCell, new ExpectedCellValue.NumericValue(41.0d)));
    assertTrue(
        AssertionExecutor.matchesCellValue(booleanCell, new ExpectedCellValue.BooleanValue(true)));
    assertFalse(
        AssertionExecutor.matchesCellValue(booleanCell, new ExpectedCellValue.BooleanValue(false)));
    assertFalse(
        AssertionExecutor.matchesCellValue(numberCell, new ExpectedCellValue.BooleanValue(true)));
    assertTrue(
        AssertionExecutor.matchesCellValue(errorCell, new ExpectedCellValue.ErrorValue("#DIV/0!")));
    assertFalse(
        AssertionExecutor.matchesCellValue(errorCell, new ExpectedCellValue.ErrorValue("#REF!")));
    assertFalse(
        AssertionExecutor.matchesCellValue(
            numberCell, new ExpectedCellValue.ErrorValue("#DIV/0!")));
    assertTrue(
        AssertionExecutor.matchesCellValue(formulaCell, new ExpectedCellValue.NumericValue(42.0d)));
    assertFalse(
        AssertionExecutor.matchesCellValue(formulaCell, new ExpectedCellValue.NumericValue(41.0d)));

    assertTrue(
        AssertionExecutor.matchesFinding(finding, finding.code(), finding.severity(), "Division"));
    assertFalse(
        AssertionExecutor.matchesFinding(
            finding, AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION, null, null));
    assertTrue(AssertionExecutor.matchesFinding(finding, finding.code(), null, null));
    assertFalse(
        AssertionExecutor.matchesFinding(
            finding, finding.code(), AnalysisSeverity.WARNING, "Division"));
    assertFalse(
        AssertionExecutor.matchesFinding(
            finding, finding.code(), finding.severity(), "Missing phrase"));
  }

  @Test
  void streamingAssertionsAndPrivateExecutionModeBranchesAreCovered() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionStepSupport stepSupport = executionStepSupport();

    GridGrindResponse.Success success =
        success(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.FULL_XSSF,
                        ExecutionModeInput.WriteMode.STREAMING_WRITE),
                    null,
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new CellMutationAction.AppendRow(
                                List.of(textCell("Owner"), textCell("Ada"))))),
                    List.of(
                        assertThat(
                            "stream-assert",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Text(
                                    "Owner")))),
                    List.of(
                        inspect(
                            "stream-read",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new InspectionQuery.GetCells())))));
    assertEquals(
        List.of("stream-assert"),
        success.assertions().stream()
            .map(dev.erst.gridgrind.contract.assertion.AssertionResult::stepId)
            .toList());

    IllegalStateException assertionModeFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                stepSupport.executeAssertionStep(
                    new AssertionStep(
                        "assert",
                        new CellSelector.ByAddress("Ops", "A1"),
                        new Assertion.CellValue(
                            new dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Text(
                                "Owner"))),
                    null,
                    new WorkbookLocation.UnsavedWorkbook(),
                    ExecutionModeInput.ReadMode.EVENT_READ));
    assertTrue(assertionModeFailure.getMessage().contains("does not support assertion steps"));

    Path workbookPath = Files.createTempFile("gridgrind-private-event-read-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.sheet("Ops").setCell("A1", dev.erst.gridgrind.excel.ExcelCellValue.text("Owner"));
      workbook.save(workbookPath);
    }
    InspectionResult.WorkbookSummaryResult eventSummary =
        assertInstanceOf(
            InspectionResult.WorkbookSummaryResult.class,
            stepSupport.executeInspectionAgainstMaterializedPath(
                inspect(
                    "event-summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary()),
                new WorkbookLocation.StoredWorkbook(workbookPath),
                ExecutionModeInput.ReadMode.EVENT_READ,
                workbookPath));
    assertEquals("event-summary", eventSummary.stepId());

    assertTrue(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                List.of(
                    inspect(
                        "workbook",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary()))),
            new ExecutionModeSelection(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF)));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                List.of()),
            new ExecutionModeSelection(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF)));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.SaveAs("copy.xlsx"),
                List.of(),
                List.of(
                    inspect(
                        "workbook",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary()))),
            new ExecutionModeSelection(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF)));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(
                    mutate(
                        new SheetSelector.ByName("Ops"), new WorkbookMutationAction.EnsureSheet())),
                List.of()),
            new ExecutionModeSelection(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF)));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                List.of(
                    inspect(
                        "workbook",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary()))),
            new ExecutionModeSelection(
                ExecutionModeInput.ReadMode.FULL_XSSF, ExecutionModeInput.WriteMode.FULL_XSSF)));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                List.of(
                    inspect(
                        "workbook",
                        new WorkbookSelector.Current(),
                        new InspectionQuery.GetWorkbookSummary()))),
            new ExecutionModeSelection(
                ExecutionModeInput.ReadMode.EVENT_READ,
                ExecutionModeInput.WriteMode.STREAMING_WRITE)));

    assertEquals(
        java.util.Optional.of("2+3"),
        ExecutionDiagnosticFields.formulaFor(new Assertion.FormulaText("2+3")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.formulaFor(new Assertion.TablePresent()));

    assertTrue(
        executor
            .executionModeFailure(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.FULL_XSSF,
                        ExecutionModeInput.WriteMode.STREAMING_WRITE),
                    null,
                    List.of(),
                    List.of(
                        assertThat(
                            "assert-without-sheet",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new Assertion.CellValue(
                                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Text(
                                    "Owner")))),
                    List.of()))
            .orElseThrow()
            .contains("requires ENSURE_SHEET before any assertion step"));
  }

  private static WorkbookPlan rewritePersistence(WorkbookPlan plan, Path workbookPath) {
    return WorkbookPlan.standard(
        plan.source(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
        plan.execution(),
        plan.formulaEnvironment(),
        plan.steps());
  }

  private static WorkbookPlan readExample(String fileName) throws IOException {
    return GridGrindJson.readRequest(Files.readAllBytes(examplesDirectory().resolve(fileName)));
  }

  private static Path examplesDirectory() {
    Path candidate = Path.of("").toAbsolutePath().normalize();
    while (candidate != null) {
      if (Files.exists(candidate.resolve("gradle.properties"))
          && Files.exists(candidate.resolve("examples"))) {
        return candidate.resolve("examples");
      }
      candidate = candidate.getParent();
    }
    throw new AssertionError("Could not locate the GridGrind examples directory.");
  }

  private static AnalysisSeverity highestSeverity(AnalysisSummaryReport summary) {
    if (summary.errorCount() > 0) {
      return AnalysisSeverity.ERROR;
    }
    if (summary.warningCount() > 0) {
      return AnalysisSeverity.WARNING;
    }
    if (summary.infoCount() > 0) {
      return AnalysisSeverity.INFO;
    }
    return null;
  }

  private static GridGrindWorkbookSurfaceReports.CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport(ExcelBorderStyle.NONE, null);
    return new GridGrindWorkbookSurfaceReports.CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }

  private static TableCellSelector.ByColumnName missingAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("BudgetTable"), "Item", textCell("Missing")),
        "Amount");
  }

  private static List<ExecutorTestPlanSupport.PendingMutation> budgetTableMutations() {
    return mutations(
        mutate(new SheetSelector.ByName("Budget"), new WorkbookMutationAction.EnsureSheet()),
        mutate(
            new dev.erst.gridgrind.contract.selector.RangeSelector.ByRange("Budget", "A1:B3"),
            new CellMutationAction.SetRange(
                List.of(
                    List.of(textCell("Item"), textCell("Amount")),
                    List.of(textCell("Hosting"), new CellInput.Numeric(100.0)),
                    List.of(textCell("Travel"), new CellInput.Numeric(50.0))))),
        mutate(
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    "BudgetTable", "Budget", "A1:B3", false, new TableStyleInput.None()))));
  }

  @Test
  void directAssertionObservationHelpersExposePresenceAndChartReads() throws IOException {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(readExecutor);
    AssertionExecutor assertionExecutor = new AssertionExecutor(readExecutor, selectorResolver);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");

      InspectionResult presence =
          assertionExecutor.presenceObservation(
              "named-range-present",
              new NamedRangeSelector.WorkbookScope("MissingBudgetTotal"),
              workbook,
              new WorkbookLocation.UnsavedWorkbook());
      InspectionResult.ChartsResult charts =
          assertionExecutor.chartsObservation(
              "charts",
              new ChartSelector.AllOnSheet("Budget"),
              workbook,
              new WorkbookLocation.UnsavedWorkbook());

      assertInstanceOf(InspectionResult.NamedRangesResult.class, presence);
      assertTrue(charts.charts().isEmpty());
      assertEquals(0, AssertionExecutor.observedCount(presence));
      assertEquals(0, AssertionExecutor.observedCount(charts));
    }
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    if (response instanceof GridGrindResponse.Failure failure) {
      fail(failure.problem().code() + ": " + failure.problem().message());
    }
    return assertInstanceOf(GridGrindResponse.Success.class, response);
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Failure.class, response);
  }

  private static ExecutionStepSupport executionStepSupport() {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(readExecutor);
    return new ExecutionStepSupport(
        new WorkbookCommandExecutor(),
        readExecutor,
        selectorResolver,
        new AssertionExecutor(readExecutor, selectorResolver),
        Files::createTempFile);
  }

  private static GridGrindResponse.Failure assertionFailure(
      DefaultGridGrindRequestExecutor executor,
      Path workbookPath,
      String stepId,
      Selector target,
      Assertion assertion) {
    return failure(
        executor.execute(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                List.of(assertThat(stepId, target, assertion)),
                List.of())));
  }
}
