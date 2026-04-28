package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

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
class GridGrindResponseEdgeCoverageTest {
  @Test
  void successFailureWorkbookSummaryAndNamedRangeBranchesAreExplicit() {
    GridGrindResponse.Success success =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), 0, false))));
    GridGrindResponse.Failure failure =
        GridGrindResponses.failure(
            GridGrindResponse.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "bad request",
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known(
                        "NEW", "NONE"))));

    assertTrue(success.warnings().isEmpty());
    assertEquals(GridGrindProtocolVersion.current(), failure.protocolVersion());
    assertEquals(
        "executionPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.PersistenceOutcome.SavedAs("budget.xlsx", " "))
            .getMessage());
    assertEquals(
        "sourcePath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.PersistenceOutcome.Overwritten(" ", "/tmp/out.xlsx"))
            .getMessage());
    assertEquals(
        "sheetCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.WorkbookSummary.Empty(-1, List.of(), 0, false))
            .getMessage());
    assertEquals(
        "namedRangeCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), -1, false))
            .getMessage());
    assertEquals(
        "sheetCount must match sheetNames size",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.WorkbookSummary.Empty(0, List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "sheetNames must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.WorkbookSummary.Empty(
                        2, List.of("Budget", "Budget"), 0, false))
            .getMessage());
    assertEquals(
        "sheetNames must not contain blank values",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.WorkbookSummary.Empty(1, List.of(" "), 0, false))
            .getMessage());
    assertEquals(
        "activeSheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), " ", List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "sheetCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        0, List.of(), "Budget", List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "activeSheetName must be present in sheetNames",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Ops", List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "selectedSheetNames must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of(), 0, false))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.NamedRangeReport.FormulaReport(
                        " ", new NamedRangeScope.Workbook(), "SUM(A1:A2)"))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.NamedRangeReport.RangeReport(
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
                    new GridGrindResponse.NamedRangeReport.RangeReport(
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
                    new GridGrindResponse.NamedRangeReport.FormulaReport(
                        "BudgetExpr", new NamedRangeScope.Workbook(), " "))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.SheetSummaryReport(
                        " ",
                        dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility.VISIBLE,
                        new GridGrindResponse.SheetProtectionReport.Unprotected(),
                        0,
                        -1,
                        -1))
            .getMessage());
  }

  @Test
  void responseDtosValidateBlankNegativeAndDuplicateBranches() {
    GridGrindResponse.CommentReport comment =
        new GridGrindResponse.CommentReport("Owner note", "Alice", true);

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
                () -> new GridGrindResponse.WindowReport(" ", "A1", 1, 1, List.of()))
            .getMessage());
    assertEquals(
        "topLeftAddress must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.WindowReport("Budget", " ", 1, 1, List.of()))
            .getMessage());
    assertEquals(
        "range must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new GridGrindResponse.MergedRegionReport(" "))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.CellHyperlinkReport(
                        " ", new HyperlinkTarget.Url("https://example.com")))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.CellCommentReport(" ", comment))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.SheetLayoutReport(
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
                () -> new GridGrindResponse.ColumnLayoutReport(-1, 8.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "widthCharacters must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.ColumnLayoutReport(0, 0.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "widthCharacters must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.ColumnLayoutReport(0, Double.NaN, false, 0, false))
            .getMessage());
    assertEquals(
        "outlineLevel must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.ColumnLayoutReport(0, 8.0d, false, -1, false))
            .getMessage());
    assertEquals(
        "rowIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.RowLayoutReport(-1, 15.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "heightPoints must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.RowLayoutReport(0, 0.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "heightPoints must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.RowLayoutReport(0, Double.NaN, false, 0, false))
            .getMessage());
    assertEquals(
        "outlineLevel must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.RowLayoutReport(0, 15.0d, false, -1, false))
            .getMessage());
    assertEquals(
        "totalFormulaCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.FormulaSurfaceReport(-1, List.of()))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetFormulaSurfaceReport(" ", 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "distinctFormulaCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetFormulaSurfaceReport("Budget", 0, -1, List.of()))
            .getMessage());
    assertEquals(
        "columnCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.WindowReport(
                        "Budget",
                        "A1",
                        1,
                        0,
                        List.of(new GridGrindResponse.WindowRowReport(0, List.of()))))
            .getMessage());
    assertEquals(
        "occurrenceCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.FormulaPatternReport("SUM(A1)", 0, List.of("A1")))
            .getMessage());
    assertEquals(
        "type must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new GridGrindResponse.TypeCountReport(" ", 1))
            .getMessage());
    assertEquals(
        "topLeftAddress must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetSchemaReport("Budget", " ", 1, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetSchemaReport(" ", "A1", 1, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "rowCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetSchemaReport("Budget", "A1", 0, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "columnCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetSchemaReport("Budget", "A1", 1, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "dataRowCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetSchemaReport("Budget", "A1", 1, 1, -1, List.of()))
            .getMessage());
    assertEquals(
        "columnAddress must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.SchemaColumnReport(
                        0, " ", "Owner", 1, 0, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "columnIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.SchemaColumnReport(
                        -1, "A", "Owner", 1, 0, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "populatedCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.SchemaColumnReport(
                        0, "A", "Owner", -1, 0, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "blankCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.SchemaColumnReport(
                        0, "A", "Owner", 1, -1, List.of(), "STRING"))
            .getMessage());
    assertEquals(
        "workbookScopedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.NamedRangeSurfaceReport(-1, 0, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "sheetScopedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.NamedRangeSurfaceReport(0, -1, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "rangeBackedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.NamedRangeSurfaceReport(0, 0, -1, 0, List.of()))
            .getMessage());
    assertEquals(
        "formulaBackedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.NamedRangeSurfaceReport(0, 0, 0, -1, List.of()))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.NamedRangeSurfaceEntryReport(
                        " ",
                        new NamedRangeScope.Workbook(),
                        "Budget!$A$1",
                        GridGrindResponse.NamedRangeBackingKind.RANGE))
            .getMessage());
    assertEquals(
        "totalCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisSummaryReport(-1, 0, 0, 0))
            .getMessage());
    assertEquals(
        "errorCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisSummaryReport(0, -1, 0, 1))
            .getMessage());
    assertEquals(
        "warningCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisSummaryReport(0, 0, -1, 1))
            .getMessage());
    assertEquals(
        "infoCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisSummaryReport(0, 0, 1, -1))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisLocationReport.Sheet(" "))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisLocationReport.Cell("Budget", " "))
            .getMessage());
    assertEquals(
        "range must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisLocationReport.Range("Budget", " "))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisLocationReport.Cell(" ", "A1"))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisLocationReport.Range(" ", "A1:B2"))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.AnalysisLocationReport.NamedRange(
                        " ", new NamedRangeScope.Workbook()))
            .getMessage());
    assertEquals(
        "title must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        " ",
                        "bad",
                        new GridGrindResponse.AnalysisLocationReport.Workbook(),
                        List.of("evidence")))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        "bad",
                        " ",
                        new GridGrindResponse.AnalysisLocationReport.Workbook(),
                        List.of("evidence")))
            .getMessage());
    assertEquals(
        "evidence must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new GridGrindResponse.AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        "bad",
                        "bad",
                        new GridGrindResponse.AnalysisLocationReport.Workbook(),
                        Arrays.asList("A1", null)))
            .getMessage());
    assertEquals(
        "checkedFormulaCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.FormulaHealthReport(
                        -1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))
            .getMessage());
    assertEquals(
        "checkedHyperlinkCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.HyperlinkHealthReport(
                        -1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))
            .getMessage());
    assertEquals(
        "checkedNamedRangeCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.NamedRangeHealthReport(
                        -1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))
            .getMessage());
  }

  @Test
  void problemContextDefaultsAndCellSubtypeTypeNamesRemainTruthful() {
    dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments parseArguments =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
            dev.erst.gridgrind.contract.dto.ProblemContext.CliArgument.named("--request"));
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

  private static GridGrindResponse.CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport(ExcelBorderStyle.NONE, null);
    return new GridGrindResponse.CellStyleReport(
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
