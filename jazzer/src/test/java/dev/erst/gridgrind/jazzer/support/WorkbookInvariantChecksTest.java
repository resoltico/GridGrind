package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.protocol.dto.AnalysisFindingCode;
import dev.erst.gridgrind.protocol.dto.AnalysisSeverity;
import dev.erst.gridgrind.protocol.dto.CellSelection;
import dev.erst.gridgrind.protocol.dto.CellAlignmentReport;
import dev.erst.gridgrind.protocol.dto.CellBorderReport;
import dev.erst.gridgrind.protocol.dto.CellBorderSideReport;
import dev.erst.gridgrind.protocol.dto.CellFillReport;
import dev.erst.gridgrind.protocol.dto.CellFontReport;
import dev.erst.gridgrind.protocol.dto.CellProtectionReport;
import dev.erst.gridgrind.protocol.dto.AutofilterEntryReport;
import dev.erst.gridgrind.protocol.dto.AutofilterHealthReport;
import dev.erst.gridgrind.protocol.dto.DataValidationEntryReport;
import dev.erst.gridgrind.protocol.dto.DataValidationHealthReport;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
import dev.erst.gridgrind.protocol.dto.FontHeightReport;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.dto.HeaderFooterTextReport;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.dto.NamedRangeScope;
import dev.erst.gridgrind.protocol.dto.NamedRangeSelection;
import dev.erst.gridgrind.protocol.dto.NamedRangeTarget;
import dev.erst.gridgrind.protocol.dto.PaneReport;
import dev.erst.gridgrind.protocol.dto.PrintAreaReport;
import dev.erst.gridgrind.protocol.dto.PrintLayoutReport;
import dev.erst.gridgrind.protocol.dto.PrintScalingReport;
import dev.erst.gridgrind.protocol.dto.PrintTitleColumnsReport;
import dev.erst.gridgrind.protocol.dto.PrintTitleRowsReport;
import dev.erst.gridgrind.protocol.dto.RangeSelection;
import dev.erst.gridgrind.protocol.dto.RichTextRunReport;
import dev.erst.gridgrind.protocol.dto.RequestWarning;
import dev.erst.gridgrind.protocol.dto.SheetProtectionSettings;
import dev.erst.gridgrind.protocol.dto.SheetSelection;
import dev.erst.gridgrind.protocol.dto.TableEntryReport;
import dev.erst.gridgrind.protocol.dto.TableHealthReport;
import dev.erst.gridgrind.protocol.dto.TableSelection;
import dev.erst.gridgrind.protocol.dto.TableStyleReport;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
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
                new RequestWarning(
                    1, "SET_CELL", "Formula references same-request sheet names with spaces.")),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
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
                            "Report",
                            List.of(
                                new RichTextRunReport(
                                    "Report",
                                    new CellFontReport(
                                        false,
                                        false,
                                        "Calibri",
                                        new FontHeightReport(220, new BigDecimal("11")),
                                        null,
                                        false,
                                        false)))))),
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
                                        "Report",
                                        null)))))),
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
                        new PaneReport.Frozen(1, 1, 1, 1),
                        125,
                        List.of(
                            new GridGrindResponse.ColumnLayoutReport(0, 12.5, false, 0, false)),
                        List.of(
                            new GridGrindResponse.RowLayoutReport(0, 18.0, false, 0, false)))),
                new WorkbookReadResult.PrintLayoutResult(
                    "print-layout",
                    new PrintLayoutReport(
                        "Budget",
                        new PrintAreaReport.Range("A1:B20"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingReport.Fit(1, 0),
                        new PrintTitleRowsReport.Band(0, 0),
                        new PrintTitleColumnsReport.Band(0, 0),
                        new HeaderFooterTextReport("Budget", "", ""),
                        new HeaderFooterTextReport("", "Page &P", ""))),
                new WorkbookReadResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Supported(
                            List.of("A2:A5"),
                                new DataValidationEntryReport.DataValidationDefinitionReport(
                                    new DataValidationRuleInput.WholeNumber(
                                    ExcelComparisonOperator.GREATER_OR_EQUAL,
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
                                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                                AnalysisSeverity.INFO,
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
            List.of(),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
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
                                AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
                                AnalysisSeverity.WARNING,
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
            List.of(),
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

  @Test
  void acceptsWorkbookShapeWithB1SheetState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.setActiveSheet("Beta");
      workbook.setSelectedSheets(List.of("Alpha", "Beta"));
      workbook.setSheetVisibility("Alpha", ExcelSheetVisibility.HIDDEN);
      workbook.setSheetProtection("Beta", protectionSettings());

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsResponseShapeWithProtectedSheetSummary(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "result.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.SheetSummaryResult(
                    "sheet",
                    new GridGrindResponse.SheetSummaryReport(
                        "Budget",
                        ExcelSheetVisibility.VERY_HIDDEN,
                        new GridGrindResponse.SheetProtectionReport.Protected(
                            protocolProtectionSettings()),
                        4,
                        7,
                        3))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsSuccessResponsesWithAutofilterAndTableReads(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "result.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(
                        new AutofilterEntryReport.SheetOwned("E1:F3"),
                        new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable"))),
                new WorkbookReadResult.TablesResult(
                    "tables",
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            "Budget",
                            "A1:C4",
                            1,
                            0,
                            List.of("Item", "Amount", "Billable"),
                            new TableStyleReport.Named(
                                "TableStyleMedium2", false, false, true, false),
                            true))),
                new WorkbookReadResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        2,
                        new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of())),
                new WorkbookReadResult.TableHealthResult(
                    "table-health",
                    new TableHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForAutofilterAndTableReads(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetAutofilters("autofilters", "Budget"),
                new WorkbookReadOperation.GetTables("tables", new TableSelection.All()),
                new WorkbookReadOperation.AnalyzeAutofilterHealth(
                    "autofilter-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeTableHealth(
                    "table-health", new TableSelection.All())));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(new AutofilterEntryReport.SheetOwned("E1:F3"))),
                new WorkbookReadResult.TablesResult(
                    "tables",
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            "Budget",
                            "A1:C4",
                            1,
                            0,
                            List.of("Item", "Amount", "Billable"),
                            new TableStyleReport.None(),
                            true))),
                new WorkbookReadResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of())),
                new WorkbookReadResult.TableHealthResult(
                    "table-health",
                    new TableHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  private static GridGrindResponse.CellStyleReport defaultStyle() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Calibri",
            new FontHeightReport(220, new BigDecimal("11")),
            null,
            false,
            false),
        new CellFillReport(ExcelFillPattern.NONE, null, null),
        new CellBorderReport(
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null)),
        new CellProtectionReport(true, false));
  }

  private static GridGrindResponse.CellReport.TextReport textCell(String address, String value) {
    return new GridGrindResponse.CellReport.TextReport(
        address, "STRING", value, defaultStyle(), null, null, value, null);
  }

  private static SheetProtectionSettings protocolProtectionSettings() {
    return new SheetProtectionSettings(
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false);
  }

  private static ExcelSheetProtectionSettings protectionSettings() {
    return new ExcelSheetProtectionSettings(
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false,
        true,
        false);
  }
}
