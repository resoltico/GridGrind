package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.protocol.CellSelection;
import dev.erst.gridgrind.protocol.DataValidationEntryReport;
import dev.erst.gridgrind.protocol.DataValidationHealthReport;
import dev.erst.gridgrind.protocol.DataValidationRuleInput;
import dev.erst.gridgrind.protocol.FontHeightReport;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.HyperlinkTarget;
import dev.erst.gridgrind.protocol.NamedRangeScope;
import dev.erst.gridgrind.protocol.NamedRangeSelection;
import dev.erst.gridgrind.protocol.NamedRangeTarget;
import dev.erst.gridgrind.protocol.RangeSelection;
import dev.erst.gridgrind.protocol.SheetSelection;
import dev.erst.gridgrind.protocol.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.WorkbookReadResult;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests for WorkbookInvariantChecks success-shape validation. */
class WorkbookInvariantChecksTest {
  @Test
  void acceptsSuccessResponsesWithOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");
    GridGrindResponse.CellStyleReport style = defaultStyle();

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "result.xlsx", workbookPath.toString()),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, false)),
                new WorkbookReadResult.NamedRangesResult(
                    "ranges",
                    List.of(
                        new GridGrindResponse.NamedRangeReport.RangeReport(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            "Budget!$B$4",
                            new NamedRangeTarget("Budget", "B4")))),
                new WorkbookReadResult.CellsResult(
                    "cells",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellReport.TextReport(
                            "A1",
                            "STRING",
                            "Report",
                            style,
                            new HyperlinkTarget.Url("https://example.com/report"),
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true),
                            "Report"))),
                new WorkbookReadResult.WindowResult(
                    "window",
                    new GridGrindResponse.WindowReport(
                        "Budget",
                        "A1",
                        1,
                        1,
                        List.of(
                            new GridGrindResponse.WindowRowReport(
                                0,
                                List.of(
                                    new GridGrindResponse.CellReport.TextReport(
                                        "A1",
                                        "STRING",
                                        "Report",
                                        style,
                                        null,
                                        null,
                                        "Report")))))),
                new WorkbookReadResult.MergedRegionsResult(
                    "merged",
                    "Budget",
                    List.of(new GridGrindResponse.MergedRegionReport("A1:B1"))),
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new WorkbookReadResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellCommentReport(
                            "A1",
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true)))),
                new WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new GridGrindResponse.SheetLayoutReport(
                        "Budget",
                        new GridGrindResponse.FreezePaneReport.Frozen(1, 1, 1, 1),
                        List.of(new GridGrindResponse.ColumnLayoutReport(0, 12.5)),
                        List.of(new GridGrindResponse.RowLayoutReport(0, 18.0)))),
                new WorkbookReadResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Supported(
                            List.of("A2:A5"),
                            new DataValidationEntryReport.DataValidationDefinitionReport(
                                new DataValidationRuleInput.WholeNumber(
                                    dev.erst.gridgrind.excel.ExcelComparisonOperator
                                        .GREATER_OR_EQUAL,
                                    "1",
                                    null),
                                true,
                                false,
                                null,
                                null)))),
                new WorkbookReadResult.FormulaSurfaceResult(
                    "formula-surface",
                    new GridGrindResponse.FormulaSurfaceReport(
                        1,
                        List.of(
                            new GridGrindResponse.SheetFormulaSurfaceReport(
                                "Budget",
                                1,
                                1,
                                List.of(
                                    new GridGrindResponse.FormulaPatternReport(
                                        "SUM(B2:B3)", 1, List.of("B4"))))))),
                new WorkbookReadResult.SheetSchemaResult(
                    "schema",
                    new GridGrindResponse.SheetSchemaReport(
                        "Budget",
                        "A1",
                        2,
                        1,
                        1,
                        List.of(
                            new GridGrindResponse.SchemaColumnReport(
                                0,
                                "A1",
                                "Item",
                                1,
                                0,
                                List.of(new GridGrindResponse.TypeCountReport("STRING", 1)),
                                "STRING")))),
                new WorkbookReadResult.NamedRangeSurfaceResult(
                    "named-range-surface",
                    new GridGrindResponse.NamedRangeSurfaceReport(
                        1,
                        0,
                        1,
                        0,
                        List.of(
                            new GridGrindResponse.NamedRangeSurfaceEntryReport(
                                "BudgetTotal",
                                new NamedRangeScope.Workbook(),
                                "Budget!$B$4",
                                GridGrindResponse.NamedRangeBackingKind.RANGE)))),
                new WorkbookReadResult.FormulaHealthResult(
                    "formula-health",
                    new GridGrindResponse.FormulaHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(1, 0, 0, 1),
                        List.of(
                            new GridGrindResponse.AnalysisFindingReport(
                                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                    .FORMULA_VOLATILE_FUNCTION,
                                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
                                "Volatile formula",
                                "Formula uses NOW().",
                                new GridGrindResponse.AnalysisLocationReport.Cell(
                                    "Budget", "B4"),
                                List.of("NOW()"))))),
                new WorkbookReadResult.DataValidationHealthResult(
                    "data-validation-health",
                    new DataValidationHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetWorkbookSummary("summary"),
                new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
                new WorkbookReadOperation.GetDataValidations(
                    "data-validations", "Budget", new RangeSelection.All()),
                new WorkbookReadOperation.GetHyperlinks(
                    "hyperlinks", "Budget", new CellSelection.AllUsedCells()),
                new WorkbookReadOperation.AnalyzeDataValidationHealth(
                    "data-validation-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                    "named-range-health", new NamedRangeSelection.All())));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, false)),
                new WorkbookReadResult.CellsResult(
                    "cells", "Budget", List.of(textCell("A1", "Report"))),
                new WorkbookReadResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Unsupported(
                            List.of("A2:A5"), "MISSING_FORMULA", "Formula is missing"))),
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new WorkbookReadResult.DataValidationHealthResult(
                    "data-validation-health",
                    new DataValidationHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(1, 1, 0, 0),
                        List.of(
                            new GridGrindResponse.AnalysisFindingReport(
                                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                    .DATA_VALIDATION_UNSUPPORTED_RULE,
                                dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.WARNING,
                                "Unsupported data-validation rule",
                                "Formula is missing",
                                new GridGrindResponse.AnalysisLocationReport.Range(
                                    "Budget", "A2:A5"),
                                List.of("A2:A5"))))),
                new WorkbookReadResult.NamedRangeHealthResult(
                    "named-range-health",
                    new GridGrindResponse.NamedRangeHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsNormalizedFileHyperlinkTargets(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.File("/tmp/report.xlsx"))))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkbookShapeWithNamedRanges() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("B4", ExcelCellValue.number(61.0));
      workbook.setNamedRange(
          new dev.erst.gridgrind.excel.ExcelNamedRangeDefinition(
              "BudgetTotal",
              new dev.erst.gridgrind.excel.ExcelNamedRangeScope.WorkbookScope(),
              new dev.erst.gridgrind.excel.ExcelNamedRangeTarget("Budget", "B4")));

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  private static GridGrindResponse.CellStyleReport defaultStyle() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        false,
        false,
        false,
        ExcelHorizontalAlignment.GENERAL,
        ExcelVerticalAlignment.BOTTOM,
        "Calibri",
        new FontHeightReport(220, new BigDecimal("11")),
        null,
        false,
        false,
        null,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE);
  }

  private static GridGrindResponse.CellReport.TextReport textCell(String address, String value) {
    return new GridGrindResponse.CellReport.TextReport(
        address, "STRING", value, defaultStyle(), null, null, value);
  }
}
