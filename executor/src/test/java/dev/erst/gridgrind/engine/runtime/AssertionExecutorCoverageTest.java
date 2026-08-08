package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.AnalysisFindingReport;
import dev.erst.gridgrind.contract.dto.AnalysisLocationReport;
import dev.erst.gridgrind.contract.dto.AnalysisSummaryReport;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAnalysisResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
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
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
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
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Coverage lock-in for Phase-4 assertion execution, helper seams, and composite failures. */
class AssertionExecutorCoverageTest extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void executesLeafAssertionsAcrossWorkbookFactsAndAnalysis() throws IOException {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Path workbookPath = Files.createTempFile("gridgrind-assertions-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    success(
        ExecutionContextFixtureSupport.execute(
            executor,
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
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
                        new CellMutationAction.SetCell(new CellInput.NumberValue(42.0d))),
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
                            new WorkbookProtectionInput(
                                true,
                                false,
                                false,
                                java.util.Optional.empty(),
                                java.util.Optional.empty()))),
                    mutate(
                        new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                        new StructuredMutationAction.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            NamedRangeTarget.range("Budget", "F2"))),
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

    WorkbookResult.Success inspected =
        success(
            ExecutionContextFixtureSupport.execute(
                executor,
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
                            allFacetCellsQuery()),
                        inspect(
                            "protection",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookProtection()),
                        inspect(
                            "sheet",
                            new SheetSelector.ByName("Budget"),
                            new SheetIntrospectionQuery.GetSheetSummary()),
                        inspect(
                            "namedRanges",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new WorkbookIntrospectionQuery.GetNamedRanges()),
                        inspect(
                            "tables",
                            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                            new WorkbookAssetIntrospectionQuery.GetTables()),
                        inspect(
                            "formulaHealth",
                            new SheetSelector.ByName("Budget"),
                            new InspectionAnalysisQuery.AnalyzeFormulaHealth())))));

    SheetInspectionResult.CellsResult cells =
        inspection(inspected, "cells", SheetInspectionResult.CellsResult.class);
    WorkbookInspectionResult.WorkbookProtectionResult protection =
        inspection(
            inspected, "protection", WorkbookInspectionResult.WorkbookProtectionResult.class);
    SheetInspectionResult.SheetSummaryResult sheet =
        inspection(inspected, "sheet", SheetInspectionResult.SheetSummaryResult.class);
    WorkbookInspectionResult.NamedRangesResult namedRanges =
        inspection(inspected, "namedRanges", WorkbookInspectionResult.NamedRangesResult.class);
    WorkbookAssetInspectionResult.TablesResult tables =
        inspection(inspected, "tables", WorkbookAssetInspectionResult.TablesResult.class);
    WorkbookAnalysisResult.FormulaHealthResult formulaHealth =
        inspection(inspected, "formulaHealth", WorkbookAnalysisResult.FormulaHealthResult.class);

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

    WorkbookResult.Success asserted =
        success(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "present-sheet-all",
                            new SheetSelector.All(),
                            new PresenceAssertion.SheetPresent()),
                        assertThat(
                            "present-sheet-by-name",
                            new SheetSelector.ByName("Budget"),
                            new PresenceAssertion.SheetPresent()),
                        assertThat(
                            "absent-sheet-by-name",
                            new SheetSelector.ByName("NonExistent"),
                            new PresenceAssertion.SheetAbsent()),
                        assertThat(
                            "present-sheet-by-names",
                            new SheetSelector.ByNames(List.of("Budget")),
                            new PresenceAssertion.SheetPresent()),
                        assertThat(
                            "present-named-range",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new PresenceAssertion.NamedRangePresent()),
                        assertThat(
                            "absent-named-range",
                            new NamedRangeSelector.WorkbookScope("MissingTotal"),
                            new PresenceAssertion.NamedRangeAbsent()),
                        assertThat(
                            "present-table",
                            new TableSelector.ByName("BudgetTable"),
                            new PresenceAssertion.TablePresent()),
                        assertThat(
                            "cell-text",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                    textValue(owner)))),
                        assertThat(
                            "cell-number",
                            new CellSelector.ByAddress("Budget", "B2"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.NumberValue(
                                    numberValue(amount)))),
                        assertThat(
                            "cell-boolean",
                            new CellSelector.ByAddress("Budget", "C2"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.BooleanValue(
                                    booleanValue(enabled)))),
                        assertThat(
                            "cell-blank",
                            new CellSelector.ByAddress("Budget", "D2"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Blank())),
                        assertThat(
                            "cell-error",
                            new CellSelector.ByAddress("Budget", "E2"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.ErrorValue(
                                    assertInstanceOf(
                                            dev.erst.gridgrind.contract.dto.CellValueReport
                                                .ErrorValue.class,
                                            evaluation(errorFormula))
                                        .errorValue()))),
                        assertThat(
                            "cell-formula-evaluation",
                            new CellSelector.ByAddress("Budget", "F2"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.NumberValue(
                                    assertInstanceOf(
                                            dev.erst.gridgrind.contract.dto.CellValueReport
                                                .NumberValue.class,
                                            evaluation(totalFormula))
                                        .numberValue()))),
                        assertThat(
                            "display-value",
                            new CellSelector.ByAddress("Budget", "B2"),
                            new CellAssertion.DisplayValue(displayValue(amount))),
                        assertThat(
                            "formula-text",
                            new CellSelector.ByAddress("Budget", "F2"),
                            new CellAssertion.FormulaText(formulaText(totalFormula))),
                        assertThat(
                            "cell-style",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new CellAssertion.CellStyle(style(owner))),
                        assertThat(
                            "workbook-protection",
                            new WorkbookSelector.Current(),
                            new WorkbookFactAssertion.WorkbookProtectionFacts(
                                protection.protection())),
                        assertThat(
                            "sheet-structure",
                            new SheetSelector.ByName("Budget"),
                            new WorkbookFactAssertion.SheetStructureFacts(sheet.sheet())),
                        assertThat(
                            "named-range-facts",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new WorkbookFactAssertion.NamedRangeFacts(namedRanges.namedRanges())),
                        assertThat(
                            "table-facts",
                            new TableSelector.ByName("BudgetTable"),
                            new WorkbookFactAssertion.TableFacts(tables.tables())),
                        assertThat(
                            "analysis-max-severity",
                            new SheetSelector.ByName("Budget"),
                            new AnalysisAssertion.AnalysisMaxSeverity(
                                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                                highestSeverity(formulaHealth.analysis().summary()))),
                        assertThat(
                            "analysis-finding-present",
                            new SheetSelector.ByName("Budget"),
                            new AnalysisAssertion.AnalysisFindingPresent(
                                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                                firstFinding.code(),
                                Optional.of(firstFinding.severity()),
                                Optional.of(firstFinding.message().substring(0, 3)))),
                        assertThat(
                            "analysis-finding-absent",
                            new SheetSelector.ByName("Budget"),
                            new AnalysisAssertion.AnalysisFindingAbsent(
                                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                                Optional.empty(),
                                Optional.empty())),
                        assertThat(
                            "all-of",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new CompositeAssertion.AllOf(
                                List.of(
                                    new CellAssertion.CellValue(
                                        new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                            textValue(owner))),
                                    new CellAssertion.DisplayValue(displayValue(owner))))),
                        assertThat(
                            "any-of",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new CompositeAssertion.AnyOf(
                                List.of(
                                    new CellAssertion.DisplayValue("Wrong"),
                                    new CellAssertion.CellValue(
                                        new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                            textValue(owner)))))),
                        assertThat(
                            "not",
                            new CellSelector.ByAddress("Budget", "A2"),
                            new CompositeAssertion.Not(
                                new CellAssertion.CellValue(
                                    new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                        "Wrong"))))),
                    List.of())));

    assertFalse(asserted.assertions().isEmpty());
    assertEquals(
        List.of(
            "present-sheet-all",
            "present-sheet-by-name",
            "absent-sheet-by-name",
            "present-sheet-by-names",
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
        ExecutionContextFixtureSupport.execute(
            executor,
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
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
                        new CellMutationAction.SetCell(new CellInput.NumberValue(42.0d))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "E2"),
                        new CellMutationAction.SetCell(formulaCell("1/0"))),
                    mutate(
                        new CellSelector.ByAddress("Budget", "F2"),
                        new CellMutationAction.SetCell(formulaCell("2+3"))),
                    mutate(
                        new WorkbookSelector.Current(),
                        new WorkbookMutationAction.SetWorkbookProtection(
                            new WorkbookProtectionInput(
                                true,
                                false,
                                false,
                                java.util.Optional.empty(),
                                java.util.Optional.empty()))),
                    mutate(
                        new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                        new StructuredMutationAction.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            NamedRangeTarget.range("Budget", "F2"))),
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

    WorkbookResult.Success inspected =
        success(
            ExecutionContextFixtureSupport.execute(
                executor,
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
                            allFacetCellsQuery()),
                        inspect(
                            "protection",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookProtection()),
                        inspect(
                            "sheet",
                            new SheetSelector.ByName("Budget"),
                            new SheetIntrospectionQuery.GetSheetSummary()),
                        inspect(
                            "namedRanges",
                            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                            new WorkbookIntrospectionQuery.GetNamedRanges()),
                        inspect(
                            "tables",
                            new TableSelector.ByName("BudgetTable"),
                            new WorkbookAssetIntrospectionQuery.GetTables()),
                        inspect(
                            "formulaHealth",
                            new SheetSelector.ByName("Budget"),
                            new InspectionAnalysisQuery.AnalyzeFormulaHealth())))));

    dev.erst.gridgrind.contract.dto.CellReport.TextReport owner =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
            inspection(inspected, "cells", SheetInspectionResult.CellsResult.class)
                .cells()
                .getFirst());
    WorkbookInspectionResult.WorkbookProtectionResult protection =
        inspection(
            inspected, "protection", WorkbookInspectionResult.WorkbookProtectionResult.class);
    SheetInspectionResult.SheetSummaryResult sheet =
        inspection(inspected, "sheet", SheetInspectionResult.SheetSummaryResult.class);
    WorkbookInspectionResult.NamedRangesResult namedRanges =
        inspection(inspected, "namedRanges", WorkbookInspectionResult.NamedRangesResult.class);
    WorkbookAssetInspectionResult.TablesResult tables =
        inspection(inspected, "tables", WorkbookAssetInspectionResult.TablesResult.class);
    WorkbookAnalysisResult.FormulaHealthResult formulaHealth =
        inspection(inspected, "formulaHealth", WorkbookAnalysisResult.FormulaHealthResult.class);
    AnalysisFindingReport firstFinding = formulaHealth.analysis().findings().getFirst();

    WorkbookResult.Failure presentMissingSheet =
        assertionFailure(
            executor,
            workbookPath,
            "present-missing-sheet",
            new SheetSelector.ByName("NonExistent"),
            new PresenceAssertion.SheetPresent());
    assertTrue(presentMissingSheet.problem().message().contains("EXPECT_SHEET_PRESENT"));
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookInspectionResult.SheetsResult.class,
                presentMissingSheet
                    .problem()
                    .assertionFailure()
                    .orElseThrow()
                    .observations()
                    .getFirst())
            .sheetNames());

    WorkbookResult.Failure absentPresentSheet =
        assertionFailure(
            executor,
            workbookPath,
            "absent-present-sheet",
            new SheetSelector.ByName("Budget"),
            new PresenceAssertion.SheetAbsent());
    assertTrue(absentPresentSheet.problem().message().contains("EXPECT_SHEET_ABSENT"));

    WorkbookResult.Failure presentMissing =
        assertionFailure(
            executor,
            workbookPath,
            "present-missing-range",
            new NamedRangeSelector.WorkbookScope("MissingTotal"),
            new PresenceAssertion.NamedRangePresent());
    assertTrue(presentMissing.problem().message().contains("EXPECT_NAMED_RANGE_PRESENT"));
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookInspectionResult.NamedRangesResult.class,
                presentMissing.problem().assertionFailure().orElseThrow().observations().getFirst())
            .namedRanges());

    WorkbookResult.Failure absentTable =
        assertionFailure(
            executor,
            workbookPath,
            "absent-table",
            new TableSelector.ByName("BudgetTable"),
            new PresenceAssertion.TableAbsent());
    assertTrue(absentTable.problem().message().contains("EXPECT_TABLE_ABSENT"));

    WorkbookResult.Failure styleMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "style-mismatch",
            new CellSelector.ByAddress("Budget", "A2"),
            new CellAssertion.CellStyle(
                new CellStyleReport(
                    "0",
                    style(owner).alignment(),
                    style(owner).font(),
                    style(owner).fill(),
                    style(owner).border(),
                    style(owner).protection())));
    assertTrue(styleMismatch.problem().message().contains("EXPECT_CELL_STYLE"));

    WorkbookResult.Failure formulaTextMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "formula-text-mismatch",
            new CellSelector.ByAddress("Budget", "F2"),
            new CellAssertion.FormulaText("1+1"));
    assertTrue(formulaTextMismatch.problem().message().contains("EXPECT_FORMULA_TEXT"));

    WorkbookProtectionReport expectedProtection = protection.protection();
    WorkbookResult.Failure protectionMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "workbook-protection-mismatch",
            new WorkbookSelector.Current(),
            new WorkbookFactAssertion.WorkbookProtectionFacts(
                new WorkbookProtectionReport(
                    !expectedProtection.structureLocked(),
                    expectedProtection.windowsLocked(),
                    expectedProtection.revisionsLocked(),
                    expectedProtection.workbookPasswordHashPresent(),
                    expectedProtection.revisionsPasswordHashPresent())));
    assertTrue(protectionMismatch.problem().message().contains("EXPECT_WORKBOOK_PROTECTION"));

    SheetSummaryReport expectedSheet = sheet.sheet();
    WorkbookResult.Failure sheetMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "sheet-structure-mismatch",
            new SheetSelector.ByName("Budget"),
            new WorkbookFactAssertion.SheetStructureFacts(
                new SheetSummaryReport(
                    expectedSheet.sheetName(),
                    expectedSheet.visibility(),
                    expectedSheet.protection(),
                    expectedSheet.physicalRowCount(),
                    expectedSheet.lastRowIndex() + 1,
                    expectedSheet.lastColumnIndex())));
    assertTrue(sheetMismatch.problem().message().contains("EXPECT_SHEET_STRUCTURE"));

    WorkbookResult.Failure namedRangeMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "named-range-facts-mismatch",
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new WorkbookFactAssertion.NamedRangeFacts(List.of()));
    assertTrue(namedRangeMismatch.problem().message().contains("EXPECT_NAMED_RANGE_FACTS"));
    assertEquals(
        namedRanges.namedRanges(),
        assertInstanceOf(
                WorkbookInspectionResult.NamedRangesResult.class,
                namedRangeMismatch
                    .problem()
                    .assertionFailure()
                    .orElseThrow()
                    .observations()
                    .getFirst())
            .namedRanges());

    WorkbookResult.Failure tableMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "table-facts-mismatch",
            new TableSelector.ByName("BudgetTable"),
            new WorkbookFactAssertion.TableFacts(List.of()));
    assertTrue(tableMismatch.problem().message().contains("EXPECT_TABLE_FACTS"));
    assertEquals(
        tables.tables(),
        assertInstanceOf(
                WorkbookAssetInspectionResult.TablesResult.class,
                tableMismatch.problem().assertionFailure().orElseThrow().observations().getFirst())
            .tables());

    WorkbookResult.Failure severityMismatch =
        assertionFailure(
            executor,
            workbookPath,
            "analysis-max-severity-mismatch",
            new SheetSelector.ByName("Budget"),
            new AnalysisAssertion.AnalysisMaxSeverity(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING));
    assertTrue(severityMismatch.problem().message().contains("EXPECT_ANALYSIS_MAX_SEVERITY"));

    WorkbookResult.Failure missingFinding =
        assertionFailure(
            executor,
            workbookPath,
            "analysis-finding-present-missing",
            new SheetSelector.ByName("Budget"),
            new AnalysisAssertion.AnalysisFindingPresent(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                Optional.empty(),
                Optional.empty()));
    assertTrue(missingFinding.problem().message().contains("missing finding"));

    WorkbookResult.Failure unexpectedFinding =
        assertionFailure(
            executor,
            workbookPath,
            "analysis-finding-absent-unexpected",
            new SheetSelector.ByName("Budget"),
            new AnalysisAssertion.AnalysisFindingAbsent(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                firstFinding.code(),
                Optional.of(firstFinding.severity()),
                Optional.of(firstFinding.message().substring(0, 3))));
    assertTrue(unexpectedFinding.problem().message().contains("unexpectedly present"));
  }

  @Test
  void compositeFailuresAndFormulaTextMismatchesReturnStructuredProblems() throws IOException {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Path workbookPath = Files.createTempFile("gridgrind-assertion-failures-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    success(
        ExecutionContextFixtureSupport.execute(
            executor,
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                List.of(
                    mutate(
                        new SheetSelector.ByName("Budget"),
                        new WorkbookMutationAction.EnsureSheet()),
                    mutate(
                        new CellSelector.ByAddress("Budget", "A1"),
                        new CellMutationAction.SetCell(textCell("Owner")))),
                List.of(),
                List.of())));

    WorkbookResult.Failure formulaMismatch =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "formula-mismatch",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CellAssertion.FormulaText("1+1"))),
                    List.of())));
    assertTrue(formulaMismatch.problem().message().contains("EXPECT_FORMULA_TEXT"));

    WorkbookResult.Failure allOfFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "all-of-failure",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CompositeAssertion.AllOf(
                                List.of(
                                    new CellAssertion.CellValue(
                                        new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                            "Owner")),
                                    new CellAssertion.DisplayValue("Wrong"))))),
                    List.of())));
    assertTrue(allOfFailure.problem().message().contains("ALL_OF failed"));

    WorkbookResult.Failure anyOfFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "any-of-failure",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CompositeAssertion.AnyOf(
                                List.of(
                                    new CellAssertion.DisplayValue("Wrong"),
                                    new CellAssertion.CellValue(
                                        new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                            "Also wrong")))))),
                    List.of())));
    assertTrue(anyOfFailure.problem().message().contains("ANY_OF failed"));

    WorkbookResult.Failure notFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "not-failure",
                            new CellSelector.ByAddress("Budget", "A1"),
                            new CompositeAssertion.Not(
                                new CellAssertion.CellValue(
                                    new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                        "Owner"))))),
                    List.of())));
    assertTrue(notFailure.problem().message().contains("NOT failed"));
    assertTrue(notFailure.problem().assertionFailure().isPresent());
  }

  @Test
  void zeroMatchDisplayFormulaAndStyleAssertionsReturnStructuredProblems() {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    WorkbookResult.Failure displayFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    List.of(
                        assertThat(
                            "display-missing-table-cell",
                            missingAmountCellTarget(),
                            new CellAssertion.DisplayValue("999"))),
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
                SheetInspectionResult.CellsResult.class,
                displayFailure.problem().assertionFailure().orElseThrow().observations().getFirst())
            .cells());

    WorkbookResult.Failure formulaFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    List.of(
                        assertThat(
                            "formula-missing-table-cell",
                            missingAmountCellTarget(),
                            new CellAssertion.FormulaText("SUM(A1:A2)"))),
                    List.of())));
    assertTrue(
        formulaFailure
            .problem()
            .message()
            .contains("EXPECT_FORMULA_TEXT resolved no matching cells"));
    assertEquals(
        "formula-missing-table-cell",
        formulaFailure.problem().assertionFailure().orElseThrow().stepId());

    WorkbookResult.Failure styleFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    List.of(
                        assertThat(
                            "style-missing-table-cell",
                            missingAmountCellTarget(),
                            new CellAssertion.CellStyle(style()))),
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

    success(
        ExecutionContextFixtureSupport.execute(
            executor, rewritePersistence(readExample("chart-request.json"), chartPath)));
    success(
        ExecutionContextFixtureSupport.execute(
            executor, rewritePersistence(readExample("pivot-request.json"), pivotPath)));

    WorkbookAssetInspectionResult.ChartsResult charts =
        inspection(
            success(
                ExecutionContextFixtureSupport.execute(
                    executor,
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(chartPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            inspect(
                                "charts",
                                new ChartSelector.AllOnSheet("Ops"),
                                new WorkbookAssetIntrospectionQuery.GetCharts()))))),
            "charts",
            WorkbookAssetInspectionResult.ChartsResult.class);
    WorkbookAssetInspectionResult.PivotTablesResult pivots =
        inspection(
            success(
                ExecutionContextFixtureSupport.execute(
                    executor,
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(pivotPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            inspect(
                                "pivots",
                                new PivotTableSelector.All(),
                                new WorkbookAssetIntrospectionQuery.GetPivotTables()))))),
            "pivots",
            WorkbookAssetInspectionResult.PivotTablesResult.class);

    WorkbookResult.Success asserted =
        success(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(chartPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "chart-facts",
                            new ChartSelector.AllOnSheet("Ops"),
                            new WorkbookFactAssertion.ChartFacts(charts.charts())),
                        assertThat(
                            "chart-present",
                            new ChartSelector.ByName("Ops", charts.charts().getFirst().name()),
                            new PresenceAssertion.ChartPresent()),
                        assertThat(
                            "chart-absent",
                            new ChartSelector.ByName("Ops", "MissingChart"),
                            new PresenceAssertion.ChartAbsent())),
                    List.of())));
    assertEquals(3, asserted.assertions().size());

    WorkbookResult.Success pivotAssertions =
        success(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(pivotPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        assertThat(
                            "pivot-facts",
                            new PivotTableSelector.All(),
                            new WorkbookFactAssertion.PivotTableFacts(pivots.pivotTables())),
                        assertThat(
                            "pivot-present",
                            new PivotTableSelector.ByName(pivots.pivotTables().getFirst().name()),
                            new PresenceAssertion.PivotTablePresent()),
                        assertThat(
                            "pivot-absent",
                            new PivotTableSelector.ByName("Missing Pivot"),
                            new PresenceAssertion.PivotTableAbsent())),
                    List.of())));
    assertEquals(3, pivotAssertions.assertions().size());

    WorkbookResult.Failure chartFactsMismatch =
        assertionFailure(
            executor,
            chartPath,
            "chart-facts-mismatch",
            new ChartSelector.AllOnSheet("Ops"),
            new WorkbookFactAssertion.ChartFacts(List.of()));
    assertTrue(chartFactsMismatch.problem().message().contains("EXPECT_CHART_FACTS"));

    WorkbookResult.Failure pivotFactsMismatch =
        assertionFailure(
            executor,
            pivotPath,
            "pivot-facts-mismatch",
            new PivotTableSelector.All(),
            new WorkbookFactAssertion.PivotTableFacts(List.of()));
    assertTrue(pivotFactsMismatch.problem().message().contains("EXPECT_PIVOT_TABLE_FACTS"));
  }

  @Test
  void analysisHelperBranchesCoverAllAnalysisFamiliesAndSeverityRanking() {
    AnalysisFindingReport finding =
        new AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_ERROR_RESULT,
            AnalysisSeverity.ERROR,
            "Formula error",
            "Division by zero",
            new AnalysisLocationReport.Cell("Budget", "E2"),
            List.of("E2"));
    AnalysisSummaryReport summary = new AnalysisSummaryReport(1, 1, 0, 0);

    List<WorkbookAnalysisResult> analyses =
        List.of(
            new WorkbookAnalysisResult.FormulaHealthResult(
                "formula", new FormulaHealthReport(1, summary, List.of(finding))),
            new WorkbookAnalysisResult.DataValidationHealthResult(
                "validation",
                new dev.erst.gridgrind.contract.dto.DataValidationHealthReport(
                    1, summary, List.of(finding))),
            new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
                "formatting",
                new dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport(
                    1, summary, List.of(finding))),
            new WorkbookAnalysisResult.AutofilterHealthResult(
                "autofilter",
                new dev.erst.gridgrind.contract.dto.AutofilterHealthReport(
                    1, summary, List.of(finding))),
            new WorkbookAnalysisResult.TableHealthResult(
                "table",
                new dev.erst.gridgrind.contract.dto.TableHealthReport(
                    1, summary, List.of(finding))),
            new WorkbookAnalysisResult.PivotTableHealthResult(
                "pivot",
                new dev.erst.gridgrind.contract.dto.PivotTableHealthReport(
                    1, summary, List.of(finding))),
            new WorkbookAnalysisResult.HyperlinkHealthResult(
                "hyperlink", new HyperlinkHealthReport(1, summary, List.of(finding))),
            new WorkbookAnalysisResult.NamedRangeHealthResult(
                "namedRange", new NamedRangeHealthReport(1, summary, List.of(finding))),
            new WorkbookAnalysisResult.WorkbookFindingsResult(
                "workbook", new WorkbookFindingsReport(summary, List.of(finding))));

    for (WorkbookAnalysisResult analysis : analyses) {
      assertEquals(summary, AssertionAnalysisEvaluator.analysisSummary(analysis));
      assertEquals(List.of(finding), AssertionAnalysisEvaluator.analysisFindings(analysis));
      assertEquals(
          java.util.Optional.of(AnalysisSeverity.ERROR),
          AssertionAnalysisEvaluator.highestSeverity(analysis));
    }

    assertEquals(-1, AssertionAnalysisEvaluator.severityRank(java.util.Optional.empty()));
    assertEquals(
        0, AssertionAnalysisEvaluator.severityRank(java.util.Optional.of(AnalysisSeverity.INFO)));
    assertEquals(
        1,
        AssertionAnalysisEvaluator.severityRank(java.util.Optional.of(AnalysisSeverity.WARNING)));
    assertEquals(
        2, AssertionAnalysisEvaluator.severityRank(java.util.Optional.of(AnalysisSeverity.ERROR)));
    assertEquals(
        java.util.Optional.of(AnalysisSeverity.WARNING),
        AssertionAnalysisEvaluator.highestSeverity(
            new WorkbookAnalysisResult.FormulaHealthResult(
                "warning",
                new FormulaHealthReport(1, new AnalysisSummaryReport(1, 0, 1, 0), List.of()))));
    assertEquals(
        java.util.Optional.of(AnalysisSeverity.INFO),
        AssertionAnalysisEvaluator.highestSeverity(
            new WorkbookAnalysisResult.FormulaHealthResult(
                "info",
                new FormulaHealthReport(1, new AnalysisSummaryReport(1, 0, 0, 1), List.of()))));
    assertEquals(
        java.util.Optional.empty(),
        AssertionAnalysisEvaluator.highestSeverity(
            new WorkbookAnalysisResult.FormulaHealthResult(
                "clean",
                new FormulaHealthReport(1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));
  }

  @Test
  void cellValueAndFindingHelpersCoverRemainingMatchingBranches() {
    AnalysisFindingReport finding =
        new AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_ERROR_RESULT,
            AnalysisSeverity.ERROR,
            "Formula error",
            "Division by zero",
            new AnalysisLocationReport.Cell("Budget", "E2"),
            List.of("E2"));
    CellStyleReport style = style();
    dev.erst.gridgrind.contract.dto.CellReport.BlankReport blankCell =
        blankReadCell("A1", "", style);
    dev.erst.gridgrind.contract.dto.CellReport.TextReport textCell =
        textReadCell("A2", "Owner", style, "Owner");
    dev.erst.gridgrind.contract.dto.CellReport.NumberReport numberCell =
        numberReadCell("B2", "42", style, 42.0d);
    dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanCell =
        booleanReadCell("C2", "TRUE", style, true);
    dev.erst.gridgrind.contract.dto.CellReport.ErrorReport errorCell =
        errorReadCell("D2", "#DIV/0!", style, "#DIV/0!");
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaCell =
        formulaReadCell(
            "E2", "42", style, "2+40", new CellValueReport.NumberValue(42.0d, Optional.empty()));

    assertTrue(AssertionValueEvaluator.matchesCellValue(blankCell, new CellScalarValue.Blank()));
    assertFalse(AssertionValueEvaluator.matchesCellValue(textCell, new CellScalarValue.Blank()));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(blankCell, new CellScalarValue.Text("Owner")));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(textCell, new CellScalarValue.Text("Owner")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(textCell, new CellScalarValue.Text("Wrong")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(textCell, new CellScalarValue.NumberValue(42.0d)));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(
            numberCell, new CellScalarValue.NumberValue(42.0d)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            numberCell, new CellScalarValue.NumberValue(41.0d)));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(
            booleanCell, new CellScalarValue.BooleanValue(true)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            booleanCell, new CellScalarValue.BooleanValue(false)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            numberCell, new CellScalarValue.BooleanValue(true)));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(
            errorCell, new CellScalarValue.ErrorValue("#DIV/0!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            errorCell, new CellScalarValue.ErrorValue("#REF!")));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            numberCell, new CellScalarValue.ErrorValue("#DIV/0!")));
    assertTrue(
        AssertionValueEvaluator.matchesCellValue(
            formulaCell, new CellScalarValue.NumberValue(42.0d)));
    assertFalse(
        AssertionValueEvaluator.matchesCellValue(
            formulaCell, new CellScalarValue.NumberValue(41.0d)));

    assertTrue(
        AssertionAnalysisEvaluator.matchesFinding(
            finding, finding.code(), Optional.of(finding.severity()), Optional.of("Division")));
    assertFalse(
        AssertionAnalysisEvaluator.matchesFinding(
            finding,
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            Optional.empty(),
            Optional.empty()));
    assertTrue(
        AssertionAnalysisEvaluator.matchesFinding(
            finding, finding.code(), Optional.empty(), Optional.empty()));
    assertFalse(
        AssertionAnalysisEvaluator.matchesFinding(
            finding,
            finding.code(),
            Optional.of(AnalysisSeverity.WARNING),
            Optional.of("Division")));
    assertFalse(
        AssertionAnalysisEvaluator.matchesFinding(
            finding,
            finding.code(),
            Optional.of(finding.severity()),
            Optional.of("Missing phrase")));
  }

  @Test
  void streamingAssertionsAndPrivateExecutionModeBranchesAreCovered() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionStepSupport stepSupport = executionStepSupport();

    WorkbookResult.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeInput.streamingWrite(),
                    null,
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new CellMutationAction.AppendRow(
                                new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                    List.of(textCell("Owner"), textCell("Ada")))))),
                    List.of(
                        assertThat(
                            "stream-assert",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                    "Owner")))),
                    List.of(
                        inspect(
                            "stream-read",
                            new CellSelector.ByAddress("Ops", "A1"),
                            allFacetCellsQuery())))));
    assertEquals(
        List.of("stream-assert"),
        success.assertions().stream()
            .map(dev.erst.gridgrind.contract.assertion.AssertionResult::stepId)
            .toList());

    WorkbookResult.Failure streamingAssertionFailure =
        failure(
            ExecutionContextFixtureSupport.execute(
                executor,
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeInput.streamingWrite(),
                    null,
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new CellMutationAction.AppendRow(
                                new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                                    List.of(textCell("Owner"), textCell("Ada")))))),
                    List.of(
                        assertThat(
                            "stream-pass",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"))),
                        assertThat(
                            "stream-fail",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                    "WrongValue")))),
                    List.of())));
    assertEquals(2, streamingAssertionFailure.assertions().size());
    assertEquals(
        dev.erst.gridgrind.contract.assertion.AssertionOutcome.PASSED,
        streamingAssertionFailure.assertions().get(0).outcome());
    assertEquals("stream-pass", streamingAssertionFailure.assertions().get(0).stepId());
    assertEquals(
        dev.erst.gridgrind.contract.assertion.AssertionOutcome.FAILED,
        streamingAssertionFailure.assertions().get(1).outcome());
    assertEquals("stream-fail", streamingAssertionFailure.assertions().get(1).stepId());

    IllegalStateException assertionModeFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                stepSupport.executeAssertionStep(
                    new AssertionStep(
                        "assert",
                        new CellSelector.ByAddress("Ops", "A1"),
                        new CellAssertion.CellValue(
                            new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"))),
                    null,
                    new WorkbookLocation.UnsavedWorkbook(),
                    ExecutionModeInput.eventRead()));
    assertTrue(assertionModeFailure.getMessage().contains("does not support assertion steps"));

    Path workbookPath = Files.createTempFile("gridgrind-private-event-read-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .sheet("Ops")
          .cells()
          .setCell("A1", dev.erst.gridgrind.excel.ExcelCellValue.text("Owner"));
      ExecutionContextFixtureSupport.saveWorkbook(workbook, workbookPath);
    }
    WorkbookInspectionResult.WorkbookSummaryResult eventSummary =
        assertInstanceOf(
            WorkbookInspectionResult.WorkbookSummaryResult.class,
            stepSupport.executeInspectionAgainstMaterializedPath(
                inspect(
                    "event-summary",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary()),
                new WorkbookLocation.StoredWorkbook(workbookPath),
                ExecutionModeInput.eventRead(),
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
                        new WorkbookIntrospectionQuery.GetWorkbookSummary()))),
            ExecutionModeInput.eventRead()));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                List.of()),
            ExecutionModeInput.eventRead()));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    "copy.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                List.of(),
                List.of(
                    inspect(
                        "workbook",
                        new WorkbookSelector.Current(),
                        new WorkbookIntrospectionQuery.GetWorkbookSummary()))),
            ExecutionModeInput.eventRead()));
    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(
                    mutate(
                        new SheetSelector.ByName("Ops"), new WorkbookMutationAction.EnsureSheet())),
                List.of()),
            ExecutionModeInput.eventRead()));
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
                        new WorkbookIntrospectionQuery.GetWorkbookSummary()))),
            ExecutionModeInput.fullXssf()));
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
                        new WorkbookIntrospectionQuery.GetWorkbookSummary()))),
            ExecutionModeInput.streamingWrite()));

    assertEquals(
        java.util.Optional.of("2+3"),
        ExecutionActionDiagnosticFields.formulaFor(new CellAssertion.FormulaText("2+3")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionActionDiagnosticFields.formulaFor(new PresenceAssertion.TablePresent()));

    assertTrue(
        executor
            .executionModeFailures(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeInput.streamingWrite(),
                    null,
                    List.of(),
                    List.of(
                        assertThat(
                            "assert-without-sheet",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                    "Owner")))),
                    List.of()))
            .stream()
            .anyMatch(
                failure -> failure.contains("requires ENSURE_SHEET before any assertion step")));
  }

  private static WorkbookPlan rewritePersistence(WorkbookPlan plan, Path workbookPath) {
    return WorkbookPlan.standard(
        plan.source(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
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

  private static CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport(ExcelBorderStyle.NONE, null);
    return new CellStyleReport(
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
                new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                    List.of(
                        List.of(textCell("Item"), textCell("Amount")),
                        List.of(textCell("Hosting"), new CellInput.NumberValue(100.0)),
                        List.of(textCell("Travel"), new CellInput.NumberValue(50.0)))))),
        mutate(
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    "BudgetTable", "Budget", "A1:B3", false, new TableStyleInput.None()))));
  }

  @Test
  void directAssertionObservationHelpersExposePresenceAndChartReads() throws IOException {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(readExecutor);
    AssertionObservationExecutor observationExecutor =
        new AssertionObservationExecutor(readExecutor, selectorResolver);

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");

      InspectionResult presence =
          observationExecutor.presenceObservation(
              "named-range-present",
              new NamedRangeSelector.WorkbookScope("MissingBudgetTotal"),
              workbook,
              new WorkbookLocation.UnsavedWorkbook());
      WorkbookAssetInspectionResult.ChartsResult charts =
          (WorkbookAssetInspectionResult.ChartsResult)
              observationExecutor.executeObservation(
                  "charts",
                  new ChartSelector.AllOnSheet("Budget"),
                  new WorkbookAssetIntrospectionQuery.GetCharts(),
                  workbook,
                  new WorkbookLocation.UnsavedWorkbook());

      assertInstanceOf(WorkbookInspectionResult.NamedRangesResult.class, presence);
      assertTrue(charts.charts().isEmpty());
      assertEquals(0, AssertionObservationExecutor.observedCount(presence));
      assertEquals(0, AssertionObservationExecutor.observedCount(charts));
    }
  }

  private static ExecutionStepSupport executionStepSupport() {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(readExecutor);
    return new ExecutionStepSupport(
        readExecutor,
        selectorResolver,
        new AssertionExecutor(readExecutor, selectorResolver),
        Files::createTempFile);
  }

  private static WorkbookResult.Failure assertionFailure(
      DefaultGridGrindRequestExecutor executor,
      Path workbookPath,
      String stepId,
      Selector target,
      Assertion assertion) {
    return failure(
        ExecutionContextFixtureSupport.execute(
            executor,
            request(
                new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of(),
                List.of(assertThat(stepId, target, assertion)),
                List.of())));
  }
}
