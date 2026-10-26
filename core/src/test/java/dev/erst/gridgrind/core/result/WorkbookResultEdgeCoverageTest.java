package dev.erst.gridgrind.core.result;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.formula.FormulaHealthReport;
import dev.erst.gridgrind.core.model.formula.FormulaPatternReport;
import dev.erst.gridgrind.core.model.formula.FormulaSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceEntryReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.SheetFormulaSurfaceReport;
import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.layout.WindowDimensionsReport;
import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.SchemaColumnReport;
import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.HyperlinkHealthReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedNameLabel;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.NamedRangeBackingKind;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Additional branch coverage for edge-case GridGrind response DTO validation. */
@SuppressWarnings("NotJavadoc")
class WorkbookResultEdgeCoverageTest {
  @Test
  void successFailureWorkbookSummaryAndNamedRangeBranchesAreExplicit() {
    RequestShape requestShape = RequestShape.known("NEW", "NONE");
    WorkbookResult.Success success =
        WorkbookResults.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult(
                    new StepId("summary"), new WorkbookSummary.Empty(0, List.of(), 0, false))));
    WorkbookResult.Failure failure =
        WorkbookResults.failure(
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "bad request",
                new dev.erst.gridgrind.core.problem.ProblemContext.ValidateRequest(requestShape)));

    assertTrue(success.warnings().isEmpty());
    assertEquals(ProtocolVersion.CURRENT, failure.protocolVersion());
    assertEquals(
        "executionPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
                        "budget.xlsx", published(" ")))
            .getMessage());
    assertEquals(
        "sourcePath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
                        Optional.of(" "), published("/tmp/out.xlsx")))
            .getMessage());
    assertEquals(
        "sheetCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookSummary.Empty(-1, List.of(), 0, false))
            .getMessage());
    assertEquals(
        "namedRangeCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookSummary.Empty(0, List.of(), -1, false))
            .getMessage());
    assertEquals(
        "sheetCount must match sheetNames size",
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookSummary.Empty(0, List.of(new SheetLabel("Budget")), 0, false))
            .getMessage());
    assertEquals(
        "sheetNames must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookSummary.Empty(
                        2, List.of(new SheetLabel("Budget"), new SheetLabel("Budget")), 0, false))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookSummary.Empty(1, List.of(new SheetLabel(" ")), 0, false))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookSummary.WithSheets(
                        1,
                        List.of(new SheetLabel("Budget")),
                        new SheetLabel(" "),
                        List.of(new SheetLabel("Budget")),
                        0,
                        false))
            .getMessage());
    assertEquals(
        "sheetCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookSummary.WithSheets(
                        0,
                        List.of(),
                        new SheetLabel("Budget"),
                        List.of(new SheetLabel("Budget")),
                        0,
                        false))
            .getMessage());
    assertEquals(
        "activeSheetName must be present in sheetNames",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookSummary.WithSheets(
                        1,
                        List.of(new SheetLabel("Budget")),
                        new SheetLabel("Ops"),
                        List.of(new SheetLabel("Budget")),
                        0,
                        false))
            .getMessage());
    assertEquals(
        "selectedSheetNames must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookSummary.WithSheets(
                        1,
                        List.of(new SheetLabel("Budget")),
                        new SheetLabel("Budget"),
                        List.of(),
                        0,
                        false))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new NamedRangeReport.FormulaReport(
                        new DefinedNameLabel(" "), new NamedRangeScope.Workbook(), "SUM(A1:A2)"))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new NamedRangeReport.RangeReport(
                        new DefinedNameLabel(" "),
                        new NamedRangeScope.Workbook(),
                        "Budget!A1",
                        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("A1"))))
            .getMessage());
    assertEquals(
        "refersToFormula must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new NamedRangeReport.RangeReport(
                        new DefinedNameLabel("BudgetTotal"),
                        new NamedRangeScope.Workbook(),
                        " ",
                        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("A1"))))
            .getMessage());
    assertEquals(
        "refersToFormula must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new NamedRangeReport.FormulaReport(
                        new DefinedNameLabel("BudgetExpr"), new NamedRangeScope.Workbook(), " "))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetSummaryReport(
                        new SheetLabel(" "),
                        dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility.VISIBLE,
                        new SheetProtectionReport.Unprotected(),
                        0,
                        -1,
                        -1))
            .getMessage());
  }

  private static WorkbookResultPersistence.PublicationOutcome.Published published(
      String executionPath) {
    return new WorkbookResultPersistence.PublicationOutcome.Published(
        executionPath,
        "0".repeat(64),
        0,
        new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification(),
        new WorkbookResultPersistence.PublicationOutcome.DurabilityEvidence
            .FileSyncedDirectoryUnestablished());
  }

  @Test
  void responseDtosValidateBlankNegativeAndDuplicateBranches() {
    CommentReport comment = new CommentReport("Owner note", "Alice", true);

    assertEquals(
        "runs must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new dev.erst.gridgrind.core.model.value.CellReport.TextReport(
                        new CellAddress("A1"),
                        Optional.of("Owner"),
                        Optional.of(style()),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of("Owner"),
                        Optional.of(List.of())))
            .getMessage());
    assertEquals(
        "runs text must concatenate to the textValue",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new dev.erst.gridgrind.core.model.value.CellReport.TextReport(
                        new CellAddress("A1"),
                        Optional.of("Owner"),
                        Optional.of(style()),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of("Owner"),
                        Optional.of(List.of(new RichTextRunReport("Mismatch", style().font())))))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WindowReport.Dense(
                        new SheetLabel(" "),
                        new CellAddress("A1"),
                        new WindowDimensionsReport(1, 1),
                        List.of()))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WindowReport.Dense(
                        new SheetLabel("Budget"),
                        new CellAddress(" "),
                        new WindowDimensionsReport(1, 1),
                        List.of()))
            .getMessage());
    assertEquals(
        "range must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new MergedRegionReport(new RangeLabel(" ")))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CellHyperlinkReport(
                        new CellAddress(" "), new HyperlinkTarget.Url("https://example.com")))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CellCommentReport(new CellAddress(" "), comment))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetLayoutReport(
                        new SheetLabel(" "),
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
                () -> new ColumnLayoutReport(-1, 8.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "widthCharacters must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ColumnLayoutReport(0, 0.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "widthCharacters must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ColumnLayoutReport(0, Double.NaN, false, 0, false))
            .getMessage());
    assertEquals(
        "outlineLevel must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ColumnLayoutReport(0, 8.0d, false, -1, false))
            .getMessage());
    assertEquals(
        "rowIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new RowLayoutReport(-1, 15.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "heightPoints must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class, () -> new RowLayoutReport(0, 0.0d, false, 0, false))
            .getMessage());
    assertEquals(
        "heightPoints must be finite and greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new RowLayoutReport(0, Double.NaN, false, 0, false))
            .getMessage());
    assertEquals(
        "outlineLevel must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new RowLayoutReport(0, 15.0d, false, -1, false))
            .getMessage());
    assertEquals(
        "totalFormulaCellCount must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new FormulaSurfaceReport(-1, List.of()))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SheetFormulaSurfaceReport(new SheetLabel(" "), 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "distinctFormulaCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SheetFormulaSurfaceReport(new SheetLabel("Budget"), 0, -1, List.of()))
            .getMessage());
    assertEquals(
        "columnCount must be greater than 0",
        assertThrows(IllegalArgumentException.class, () -> new WindowDimensionsReport(1, 0))
            .getMessage());
    assertEquals(
        "occurrenceCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new FormulaPatternReport("SUM(A1)", 0, List.of(new CellAddress("A1"))))
            .getMessage());
    assertEquals(
        "type must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new TypeCountReport(" ", 1))
            .getMessage());
    assertEquals(
        "type must be one of TEXT, NUMBER, BOOLEAN, ERROR, DATE, TIME, DATE_TIME but was STRING",
        assertThrows(IllegalArgumentException.class, () -> new TypeCountReport("STRING", 1))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetSchemaReport(
                        new SheetLabel("Budget"), new CellAddress(" "), 1, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetSchemaReport(
                        new SheetLabel(" "), new CellAddress("A1"), 1, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "rowCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetSchemaReport(
                        new SheetLabel("Budget"), new CellAddress("A1"), 0, 1, 0, List.of()))
            .getMessage());
    assertEquals(
        "columnCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetSchemaReport(
                        new SheetLabel("Budget"), new CellAddress("A1"), 1, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "dataRowCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetSchemaReport(
                        new SheetLabel("Budget"), new CellAddress("A1"), 1, 1, -1, List.of()))
            .getMessage());
    assertEquals(
        "columnAddress must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SchemaColumnReport(0, " ", "Owner", 1, 0, List.of(), "TEXT"))
            .getMessage());
    assertEquals(
        "columnIndex must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SchemaColumnReport(-1, "A", "Owner", 1, 0, List.of(), "TEXT"))
            .getMessage());
    assertEquals(
        "populatedCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SchemaColumnReport(0, "A", "Owner", -1, 0, List.of(), "TEXT"))
            .getMessage());
    assertEquals(
        "blankCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SchemaColumnReport(0, "A", "Owner", 1, -1, List.of(), "TEXT"))
            .getMessage());
    assertEquals(
        "observedTypes counts must sum to populatedCellCount",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SchemaColumnReport(
                        0, "A", "Owner", 2, 0, List.of(new TypeCountReport("TEXT", 1)), "TEXT"))
            .getMessage());
    assertEquals(
        "observedTypes must not contain duplicate type TEXT",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SchemaColumnReport(
                        0,
                        "A",
                        "Owner",
                        2,
                        0,
                        List.of(new TypeCountReport("TEXT", 1), new TypeCountReport("TEXT", 1)),
                        "TEXT"))
            .getMessage());
    assertEquals(
        "dominantType must be omitted or match one observedTypes entry",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SchemaColumnReport(
                        0, "A", "Owner", 1, 0, List.of(new TypeCountReport("TEXT", 1)), "NUMBER"))
            .getMessage());
    assertEquals(
        "workbookScopedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new NamedRangeSurfaceReport(-1, 0, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "sheetScopedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new NamedRangeSurfaceReport(0, -1, 0, 0, List.of()))
            .getMessage());
    assertEquals(
        "rangeBackedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new NamedRangeSurfaceReport(0, 0, -1, 0, List.of()))
            .getMessage());
    assertEquals(
        "formulaBackedCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new NamedRangeSurfaceReport(0, 0, 0, -1, List.of()))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new NamedRangeSurfaceEntryReport(
                        " ",
                        new NamedRangeScope.Workbook(),
                        "Budget!$A$1",
                        NamedRangeBackingKind.RANGE))
            .getMessage());
    assertEquals(
        "totalCount must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new AnalysisSummaryReport(-1, 0, 0, 0))
            .getMessage());
    assertEquals(
        "errorCount must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new AnalysisSummaryReport(0, -1, 0, 1))
            .getMessage());
    assertEquals(
        "warningCount must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new AnalysisSummaryReport(0, 0, -1, 1))
            .getMessage());
    assertEquals(
        "infoCount must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new AnalysisSummaryReport(0, 0, 1, -1))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AnalysisLocationReport.Sheet(new SheetLabel(" ")))
            .getMessage());
    assertEquals(
        "address must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AnalysisLocationReport.Cell(new SheetLabel("Budget"), " "))
            .getMessage());
    assertEquals(
        "range must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AnalysisLocationReport.Range(new SheetLabel("Budget"), " "))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AnalysisLocationReport.Cell(new SheetLabel(" "), "A1"))
            .getMessage());
    assertEquals(
        "sheetName must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AnalysisLocationReport.Range(new SheetLabel(" "), "A1:B2"))
            .getMessage());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new AnalysisLocationReport.NamedRange(" ", new NamedRangeScope.Workbook()))
            .getMessage());
    assertEquals(
        "title must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        " ",
                        "bad",
                        new AnalysisLocationReport.Workbook(),
                        List.of("evidence")))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        "bad",
                        " ",
                        new AnalysisLocationReport.Workbook(),
                        List.of("evidence")))
            .getMessage());
    assertEquals(
        "evidence must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new AnalysisFindingReport(
                        AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
                        AnalysisSeverity.ERROR,
                        "bad",
                        "bad",
                        new AnalysisLocationReport.Workbook(),
                        Arrays.asList("A1", null)))
            .getMessage());
    assertEquals(
        "checkedFormulaCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new FormulaHealthReport(-1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))
            .getMessage());
    assertEquals(
        "checkedHyperlinkCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new HyperlinkHealthReport(-1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))
            .getMessage());
    assertEquals(
        "checkedNamedRangeCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new NamedRangeHealthReport(
                        -1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))
            .getMessage());
  }

  @Test
  void problemContextDefaultsAndCellSubtypeTypeNamesRemainTruthful() {
    dev.erst.gridgrind.core.problem.ProblemContext.ParseArguments parseArguments =
        new dev.erst.gridgrind.core.problem.ProblemContext.ParseArguments(
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.CliArgument.named(
                "--request"));
    dev.erst.gridgrind.core.model.value.CellReport.BooleanReport booleanCell =
        new dev.erst.gridgrind.core.model.value.CellReport.BooleanReport(
            new CellAddress("A2"),
            java.util.Optional.of("TRUE"),
            java.util.Optional.of(style()),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of(true));
    dev.erst.gridgrind.core.model.value.CellReport.ErrorReport errorCell =
        new dev.erst.gridgrind.core.model.value.CellReport.ErrorReport(
            new CellAddress("A3"),
            java.util.Optional.of("#DIV/0!"),
            java.util.Optional.of(style()),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of("#DIV/0!"));

    assertEquals("PARSE_ARGUMENTS", parseArguments.type());
    assertEquals(java.util.Optional.of("--request"), parseArguments.argumentName());
    assertEquals("BOOLEAN", booleanCell.type());
    assertEquals("ERROR", errorCell.type());
  }

  private static CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
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
}
