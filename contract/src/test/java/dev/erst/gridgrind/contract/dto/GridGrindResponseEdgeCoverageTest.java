package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Additional branch coverage for edge-case GridGrind response DTO validation. */
@SuppressWarnings("NotJavadoc")
class GridGrindResponseEdgeCoverageTest {
  @Test
  void successFailureWorkbookSummaryAndNamedRangeBranchesAreExplicit() {
    RequestShape requestShape = RequestShape.known("NEW", "NONE");
    GridGrindResponse.Success success =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty(
                        0, List.of(), 0, false))));
    GridGrindResponse.Failure failure =
        GridGrindResponses.failure(
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "bad request",
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(requestShape)));

    assertTrue(success.warnings().isEmpty());
    assertEquals(GridGrindProtocolVersion.current(), failure.protocolVersion());
    assertEquals(
        "executionPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponsePersistence.PersistenceOutcome.SavedAs("budget.xlsx", " "))
            .getMessage());
    assertEquals(
        "sourcePath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponsePersistence.PersistenceOutcome.Overwritten(
                        " ", "/tmp/out.xlsx"))
            .getMessage());
    assertEquals(
        "sheetCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty(
                        -1, List.of(), 0, false))
            .getMessage());
    assertEquals(
        "namedRangeCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty(
                        0, List.of(), -1, false))
            .getMessage());
    assertEquals(
        "sheetCount must match sheetNames size",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty(
                        0, List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "sheetNames must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty(
                        2, List.of("Budget", "Budget"), 0, false))
            .getMessage());
    assertEquals(
        "sheetNames must not contain blank values",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.Empty(
                        1, List.of(" "), 0, false))
            .getMessage());
    assertEquals(
        "activeSheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), " ", List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "sheetCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets(
                        0, List.of(), "Budget", List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "activeSheetName must be present in sheetNames",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Ops", List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "selectedSheetNames must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of(), 0, false))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.NamedRangeReport.FormulaReport(
                        " ", new NamedRangeScope.Workbook(), "SUM(A1:A2)"))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.NamedRangeReport.RangeReport(
                        " ",
                        new NamedRangeScope.Workbook(),
                        "Budget!A1",
                        new NamedRangeTarget("Budget", "A1")))
            .getMessage());
    assertEquals(
        "refersToFormula must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.NamedRangeReport.RangeReport(
                        "BudgetTotal",
                        new NamedRangeScope.Workbook(),
                        " ",
                        new NamedRangeTarget("Budget", "A1")))
            .getMessage());
    assertEquals(
        "refersToFormula must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.NamedRangeReport.FormulaReport(
                        "BudgetExpr", new NamedRangeScope.Workbook(), " "))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindWorkbookSurfaceReports.SheetSummaryReport(
                        " ",
                        dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VISIBLE,
                        new GridGrindWorkbookSurfaceReports.SheetProtectionReport.Unprotected(),
                        0,
                        -1,
                        -1))
            .getMessage());
  }

  @Test
  void responseDtosValidateBlankNegativeAndDuplicateBranches() {
    GridGrindWorkbookSurfaceReports.CommentReport comment =
        new GridGrindWorkbookSurfaceReports.CommentReport("Owner note", "Alice", true);

    assertEquals(
        "richText must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
                        "A1",
                        "STRING",
                        "Owner",
                        style(),
                        Optional.empty(),
                        Optional.empty(),
                        "Owner",
                        Optional.of(List.of())))
            .getMessage());
    assertEquals(
        "richText run text must concatenate to the stringValue",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
                        "A1",
                        "STRING",
                        "Owner",
                        style(),
                        Optional.empty(),
                        Optional.empty(),
                        "Owner",
                        Optional.of(List.of(new RichTextRunReport("Mismatch", style().font())))))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindLayoutSurfaceReports.WindowReport(" ", "A1", 1, 1, List.of()))
            .getMessage());
    assertEquals(
        "topLeftAddress must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.WindowReport("Budget", " ", 1, 1, List.of()))
            .getMessage());
    assertEquals(
        "range must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindLayoutSurfaceReports.MergedRegionReport(" "))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.CellHyperlinkReport(
                        " ", new HyperlinkTarget.Url("https://example.com")))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindLayoutSurfaceReports.CellCommentReport(" ", comment))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.SheetLayoutReport(
                        " ",
                        new PaneReport.None(),
                        100,
                        SheetPresentationReport.defaults(),
                        List.of(),
                        List.of()))
            .getMessage());
    assertEquals(
        "columnIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.ColumnLayoutReport(-1, 8.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "widthCharacters must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.ColumnLayoutReport(0, 0.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "widthCharacters must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.ColumnLayoutReport(
                        0, Double.NaN, false, 0, false))
            .getMessage());
    assertEquals(
        "outlineLevel must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.ColumnLayoutReport(0, 8.0d, false, -1, false))
            .getMessage());
    assertEquals(
        "rowIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindLayoutSurfaceReports.RowLayoutReport(-1, 15.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "heightPoints must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindLayoutSurfaceReports.RowLayoutReport(0, 0.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "heightPoints must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.RowLayoutReport(
                        0, Double.NaN, false, 0, false))
            .getMessage());
    assertEquals(
        "outlineLevel must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindLayoutSurfaceReports.RowLayoutReport(0, 15.0d, false, -1, false))
            .getMessage());
    assertEquals(
        "totalFormulaCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindSchemaAndFormulaReports.FormulaSurfaceReport(-1, List.of()))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SheetFormulaSurfaceReport(
                        " ", 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "distinctFormulaCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SheetFormulaSurfaceReport(
                        "Budget", 0, -1, List.of()))
            .getMessage());
    assertEquals(
        "columnCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindLayoutSurfaceReports.WindowReport(
                        "Budget",
                        "A1",
                        1,
                        0,
                        List.of(new GridGrindLayoutSurfaceReports.WindowRowReport(0, List.of()))))
            .getMessage());
    assertEquals(
        "occurrenceCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.FormulaPatternReport(
                        "SUM(A1)", 0, List.of("A1")))
            .getMessage());
    assertEquals(
        "type must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindSchemaAndFormulaReports.TypeCountReport(" ", 1))
            .getMessage());
    assertEquals(
        "topLeftAddress must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SheetSchemaReport(
                        "Budget", " ", 1, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SheetSchemaReport(
                        " ", "A1", 1, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "rowCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SheetSchemaReport(
                        "Budget", "A1", 0, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "columnCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SheetSchemaReport(
                        "Budget", "A1", 1, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "dataRowCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SheetSchemaReport(
                        "Budget", "A1", 1, 1, -1, List.of()))
            .getMessage());
    assertEquals(
        "columnAddress must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SchemaColumnReport(
                        0, " ", "Owner", 1, 0, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "columnIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SchemaColumnReport(
                        -1, "A", "Owner", 1, 0, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "populatedCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SchemaColumnReport(
                        0, "A", "Owner", -1, 0, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "blankCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.SchemaColumnReport(
                        0, "A", "Owner", 1, -1, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "workbookScopedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.NamedRangeSurfaceReport(
                        -1, 0, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "sheetScopedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.NamedRangeSurfaceReport(
                        0, -1, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "rangeBackedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.NamedRangeSurfaceReport(
                        0, 0, -1, 0, List.of()))
            .getMessage());
    assertEquals(
        "formulaBackedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.NamedRangeSurfaceReport(
                        0, 0, 0, -1, List.of()))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindSchemaAndFormulaReports.NamedRangeSurfaceEntryReport(
                        " ",
                        new NamedRangeScope.Workbook(),
                        "Budget!$A$1",
                        GridGrindSchemaAndFormulaReports.NamedRangeBackingKind.RANGE))
            .getMessage());
    assertEquals(
        "totalCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisSummaryReport(-1, 0, 0, 0))
            .getMessage());
    assertEquals(
        "errorCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisSummaryReport(0, -1, 0, 1))
            .getMessage());
    assertEquals(
        "warningCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisSummaryReport(0, 0, -1, 1))
            .getMessage());
    assertEquals(
        "infoCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisSummaryReport(0, 0, 1, -1))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisLocationReport.Sheet(" "))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisLocationReport.Cell("Budget", " "))
            .getMessage());
    assertEquals(
        "range must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisLocationReport.Range("Budget", " "))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisLocationReport.Cell(" ", "A1"))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindAnalysisReports.AnalysisLocationReport.Range(" ", "A1:B2"))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindAnalysisReports.AnalysisLocationReport.NamedRange(
                        " ", new NamedRangeScope.Workbook()))
            .getMessage());
    assertEquals(
        "title must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindAnalysisReports.AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        " ",
                        "bad",
                        new GridGrindAnalysisReports.AnalysisLocationReport.Workbook(),
                        List.of("evidence")))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindAnalysisReports.AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        "bad",
                        " ",
                        new GridGrindAnalysisReports.AnalysisLocationReport.Workbook(),
                        List.of("evidence")))
            .getMessage());
    assertEquals(
        "evidence must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new GridGrindAnalysisReports.AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        "bad",
                        "bad",
                        new GridGrindAnalysisReports.AnalysisLocationReport.Workbook(),
                        Arrays.asList("A1", null)))
            .getMessage());
    assertEquals(
        "checkedFormulaCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindAnalysisReports.FormulaHealthReport(
                        -1,
                        new GridGrindAnalysisReports.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))
            .getMessage());
    assertEquals(
        "checkedHyperlinkCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindAnalysisReports.HyperlinkHealthReport(
                        -1,
                        new GridGrindAnalysisReports.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))
            .getMessage());
    assertEquals(
        "checkedNamedRangeCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindAnalysisReports.NamedRangeHealthReport(
                        -1,
                        new GridGrindAnalysisReports.AnalysisSummaryReport(0, 0, 0, 0),
                        List.of()))
            .getMessage());
  }

  @Test
  void problemContextDefaultsAndCellSubtypeTypeNamesRemainTruthful() {
    dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments parseArguments =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                "--request"));
    dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BooleanReport(
            "A2",
            "BOOLEAN",
            "TRUE",
            style(),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            true);
    dev.erst.gridgrind.contract.dto.CellReport.ErrorReport errorCell =
        new dev.erst.gridgrind.contract.dto.CellReport.ErrorReport(
            "A3",
            "ERROR",
            "#DIV/0!",
            style(),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            "#DIV/0!");

    assertEquals("PARSE_ARGUMENTS", parseArguments.stage());
    assertEquals(java.util.Optional.of("--request"), parseArguments.argumentName());
    assertEquals("BOOLEAN", booleanCell.effectiveType());
    assertEquals("ERROR", errorCell.effectiveType());
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
}
