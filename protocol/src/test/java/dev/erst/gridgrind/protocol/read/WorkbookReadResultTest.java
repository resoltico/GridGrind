package dev.erst.gridgrind.protocol.read;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.dto.*;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for workbook-read result invariants. */
class WorkbookReadResultTest {
  @Test
  void packageSecurityResultRejectsBlankRequestIdAndNullSecurity() {
    WorkbookReadResult.PackageSecurityResult result =
        new WorkbookReadResult.PackageSecurityResult(
            "security",
            new OoxmlPackageSecurityReport(
                new OoxmlEncryptionReport(false, null, null, null, null, null, null, null),
                List.of()));

    assertFalse(result.security().encryption().encrypted());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.PackageSecurityResult(
                " ",
                new OoxmlPackageSecurityReport(
                    new OoxmlEncryptionReport(false, null, null, null, null, null, null, null),
                    List.of())));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.PackageSecurityResult("security", null));
  }

  @Test
  void workbookSummaryResultRejectsBlankRequestId() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookReadResult.WorkbookSummaryResult(
                    " ",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 0, false)));

    assertEquals("requestId must not be blank", exception.getMessage());
  }

  @Test
  void cellsResultCopiesCells() {
    WorkbookReadResult.CellsResult result =
        new WorkbookReadResult.CellsResult(
            "cells",
            "Budget",
            List.of(
                new GridGrindResponse.CellReport.BlankReport(
                    "A1", "BLANK", "", defaultStyle(), null, null)));

    assertEquals("A1", result.cells().getFirst().address());
  }

  @Test
  void analysisReadResultsRejectBlankRequestIdAndNullAnalysis() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 0, 0, 1);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            AnalysisSeverity.INFO,
            "Volatile formula",
            "Formula uses NOW().",
            new GridGrindResponse.AnalysisLocationReport.Workbook(),
            List.of("NOW()"));
    GridGrindResponse.HyperlinkHealthReport hyperlinkHealth =
        new GridGrindResponse.HyperlinkHealthReport(1, summary, List.of(finding));
    GridGrindResponse.NamedRangeHealthReport namedRangeHealth =
        new GridGrindResponse.NamedRangeHealthReport(1, summary, List.of(finding));
    GridGrindResponse.WorkbookFindingsReport workbookFindings =
        new GridGrindResponse.WorkbookFindingsReport(summary, List.of(finding));

    assertEquals(
        1,
        new WorkbookReadResult.HyperlinkHealthResult("hyperlink-health", hyperlinkHealth)
            .analysis()
            .checkedHyperlinkCount());
    assertEquals(
        1,
        new WorkbookReadResult.NamedRangeHealthResult("named-range-health", namedRangeHealth)
            .analysis()
            .checkedNamedRangeCount());
    assertEquals(
        1,
        new WorkbookReadResult.WorkbookFindingsResult("workbook-findings", workbookFindings)
            .analysis()
            .summary()
            .totalCount());

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.HyperlinkHealthResult(" ", hyperlinkHealth));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.NamedRangeHealthResult("named-range-health", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.WorkbookFindingsResult(" ", workbookFindings));
  }

  @Test
  void pivotReadResultsCopyEntriesAndRejectInvalidState() {
    WorkbookReadResult.PivotTablesResult pivotTables =
        new WorkbookReadResult.PivotTablesResult(
            "pivots",
            List.of(
                new PivotTableReport.Supported(
                    "Sales Pivot 2026",
                    "Report",
                    new PivotTableReport.Anchor("C5", "C5:G9"),
                    new PivotTableReport.Source.Range("Data", "A1:D5"),
                    List.of(new PivotTableReport.Field(0, "Region")),
                    List.of(new PivotTableReport.Field(1, "Stage")),
                    List.of(new PivotTableReport.Field(2, "Owner")),
                    List.of(
                        new PivotTableReport.DataField(
                            3,
                            "Amount",
                            dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
                            "Total Amount",
                            "#,##0.00")),
                    true)));
    PivotTableHealthReport health =
        new PivotTableHealthReport(
            1,
            new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0),
            List.of(
                new GridGrindResponse.AnalysisFindingReport(
                    AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
                    AnalysisSeverity.WARNING,
                    "Pivot table name is missing",
                    "GridGrind assigned a synthetic identifier for readback.",
                    new GridGrindResponse.AnalysisLocationReport.Sheet("Report"),
                    List.of("_GG_PIVOT_Report_A3"))));
    WorkbookReadResult.PivotTableHealthResult pivotTableHealth =
        new WorkbookReadResult.PivotTableHealthResult("pivot-health", health);

    assertEquals("Sales Pivot 2026", pivotTables.pivotTables().getFirst().name());
    assertEquals(1, pivotTableHealth.analysis().checkedPivotTableCount());
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadResult.PivotTablesResult("pivots", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.PivotTableHealthResult("pivot-health", null));
  }

  @Test
  void dataValidationResultsCopyEntriesAndRejectInvalidState() {
    WorkbookReadResult.DataValidationsResult result =
        new WorkbookReadResult.DataValidationsResult(
            "validations",
            "Budget",
            List.of(
                new DataValidationEntryReport.Supported(
                    List.of("A2:A5"),
                    new DataValidationEntryReport.DataValidationDefinitionReport(
                        new DataValidationRuleInput.WholeNumber(
                            ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                        true,
                        false,
                        new DataValidationPromptInput("Priority", "Use 1 or greater.", true),
                        null))));

    assertEquals("A2:A5", result.validations().getFirst().ranges().getFirst());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.DataValidationsResult(" ", "Budget", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.DataValidationsResult("validations", "Budget", null));
  }

  @Test
  void dataValidationHealthResultsRejectBlankRequestIdAndNullAnalysis() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(2, 1, 1, 0);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES,
            AnalysisSeverity.WARNING,
            "Overlapping data validations",
            "Rules overlap on the same cells.",
            new GridGrindResponse.AnalysisLocationReport.Range("Budget", "A3:A4"),
            List.of("A1:A5", "A3:A4"));
    DataValidationHealthReport health =
        new DataValidationHealthReport(2, summary, List.of(finding));

    assertEquals(
        2,
        new WorkbookReadResult.DataValidationHealthResult("validation-health", health)
            .analysis()
            .checkedValidationCount());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.DataValidationHealthResult(" ", health));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.DataValidationHealthResult("validation-health", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationHealthReport(-1, summary, List.of(finding)));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationHealthReport(1, null, List.of(finding)));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationHealthReport(1, summary, List.of(finding, null)));
  }

  @Test
  void conditionalFormattingResultsCopyEntriesAndRejectInvalidState() {
    WorkbookReadResult.ConditionalFormattingResult result =
        new WorkbookReadResult.ConditionalFormattingResult(
            "conditional-formatting",
            "Budget",
            List.of(
                new ConditionalFormattingEntryReport(
                    List.of("A2:A5"),
                    List.of(
                        new ConditionalFormattingRuleReport.FormulaRule(
                            1,
                            true,
                            "A2>0",
                            new DifferentialStyleReport(
                                "0.00", true, null, null, null, null, null, null, null,
                                List.of()))))));

    assertEquals("A2:A5", result.conditionalFormattingBlocks().getFirst().ranges().getFirst());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.ConditionalFormattingResult(" ", "Budget", List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadResult.ConditionalFormattingResult(
                "conditional-formatting", "Budget", null));
  }

  @Test
  void conditionalFormattingHealthResultsRejectBlankRequestIdAndNullAnalysis() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 1, 0, 0);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
            AnalysisSeverity.ERROR,
            "Priority collision",
            "Conditional-formatting priorities collide.",
            new GridGrindResponse.AnalysisLocationReport.Sheet("Budget"),
            List.of("FORMULA_RULE@Budget!A1:A3"));
    ConditionalFormattingHealthReport health =
        new ConditionalFormattingHealthReport(1, summary, List.of(finding));

    assertEquals(
        1,
        new WorkbookReadResult.ConditionalFormattingHealthResult(
                "conditional-formatting-health", health)
            .analysis()
            .checkedConditionalFormattingBlockCount());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.ConditionalFormattingHealthResult(" ", health));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadResult.ConditionalFormattingHealthResult(
                "conditional-formatting-health", null));
  }

  @Test
  void drawingReadResultsCopyEntriesAndRejectInvalidState() {
    DrawingAnchorReport.TwoCell pictureAnchor =
        new DrawingAnchorReport.TwoCell(
            new DrawingMarkerReport(1, 2, 0, 0),
            new DrawingMarkerReport(4, 6, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    DrawingAnchorReport.OneCell shapeAnchor =
        new DrawingAnchorReport.OneCell(new DrawingMarkerReport(3, 4, 0, 0), 10L, 20L, null);
    DrawingObjectReport.Picture picture =
        new DrawingObjectReport.Picture(
            "OpsPicture",
            pictureAnchor,
            ExcelPictureFormat.PNG,
            "image/png",
            68L,
            "abc123",
            null,
            null,
            "Queue preview");
    DrawingObjectReport.Shape shape =
        new DrawingObjectReport.Shape(
            "OpsShape", shapeAnchor, ExcelDrawingShapeKind.SIMPLE_SHAPE, "rect", "Queue", 0);
    WorkbookReadResult.DrawingObjectsResult drawingObjects =
        new WorkbookReadResult.DrawingObjectsResult("drawing", "Budget", List.of(picture, shape));
    WorkbookReadResult.DrawingObjectPayloadResult drawingPayload =
        new WorkbookReadResult.DrawingObjectPayloadResult(
            "payload",
            "Budget",
            new DrawingObjectPayloadReport.EmbeddedObject(
                "OpsEmbed",
                ExcelEmbeddedObjectPackagingKind.RAW_PACKAGE,
                "application/octet-stream",
                "payload.txt",
                "def456",
                "cGF5bG9hZA==",
                "Payload",
                "payload.txt"));

    assertEquals("OpsPicture", drawingObjects.drawingObjects().getFirst().name());
    assertEquals("OpsEmbed", drawingPayload.payload().name());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.DrawingObjectsResult(" ", "Budget", List.of(picture)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.DrawingObjectsResult("drawing", " ", List.of(picture)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.DrawingObjectsResult("drawing", "Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.DrawingObjectPayloadResult(
                " ", "Budget", drawingPayload.payload()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.DrawingObjectPayloadResult(
                "payload", " ", drawingPayload.payload()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.DrawingObjectPayloadResult("payload", "Budget", null));
  }

  @Test
  void chartReadResultsCopyEntriesAndRejectInvalidState() {
    DrawingAnchorReport.TwoCell anchor =
        new DrawingAnchorReport.TwoCell(
            new DrawingMarkerReport(1, 2, 0, 0),
            new DrawingMarkerReport(6, 12, 0, 0),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ChartReport.Bar chart =
        new ChartReport.Bar(
            "OpsChart",
            anchor,
            new ChartReport.Title.Text("Roadmap"),
            new ChartReport.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            true,
            ExcelChartBarDirection.COLUMN,
            List.of(
                new ChartReport.Axis(
                    ExcelChartAxisKind.CATEGORY,
                    ExcelChartAxisPosition.BOTTOM,
                    ExcelChartAxisCrosses.AUTO_ZERO,
                    true),
                new ChartReport.Axis(
                    ExcelChartAxisKind.VALUE,
                    ExcelChartAxisPosition.LEFT,
                    ExcelChartAxisCrosses.AUTO_ZERO,
                    true)),
            List.of(
                new ChartReport.Series(
                    new ChartReport.Title.Formula("Chart!$B$1", "Plan"),
                    new ChartReport.DataSource.StringReference("ChartCategories", List.of("Jan")),
                    new ChartReport.DataSource.NumericReference(
                        "ChartValues", null, List.of("10.0")))));
    WorkbookReadResult.ChartsResult charts =
        new WorkbookReadResult.ChartsResult("charts", "Budget", List.of(chart));

    assertEquals("OpsChart", charts.charts().getFirst().name());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.ChartsResult(" ", "Budget", List.of(chart)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.ChartsResult("charts", " ", List.of(chart)));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.ChartsResult("charts", "Budget", null));
  }

  @Test
  void autofilterAndTableResultsCopyEntriesAndRejectInvalidState() {
    WorkbookReadResult.AutofiltersResult autofilters =
        new WorkbookReadResult.AutofiltersResult(
            "filters",
            "Budget",
            List.of(
                new AutofilterEntryReport.SheetOwned("E1:F4"),
                new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable")));
    WorkbookReadResult.TablesResult tables =
        new WorkbookReadResult.TablesResult(
            "tables",
            List.of(
                new TableEntryReport(
                    "BudgetTable",
                    "Budget",
                    "A1:C4",
                    1,
                    1,
                    List.of("Item", "Amount", "Billable"),
                    new TableStyleReport.Named("TableStyleMedium2", false, false, true, false),
                    true)));

    assertEquals("E1:F4", autofilters.autofilters().getFirst().range());
    assertEquals("BudgetTable", tables.tables().getFirst().name());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.AutofiltersResult(" ", "Budget", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.AutofiltersResult("filters", "Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadResult.TablesResult("tables", null));
  }

  @Test
  void autofilterAndTableHealthResultsRejectBlankRequestIdAndNullAnalysis() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
            AnalysisSeverity.WARNING,
            "Table autofilter does not match table range",
            "Table-owned autofilter range must match the table range excluding any totals row.",
            new GridGrindResponse.AnalysisLocationReport.Range("Budget", "A1:C3"),
            List.of("BudgetTable", "A1:C4", "A1:C3"));
    AutofilterHealthReport autofilterHealth =
        new AutofilterHealthReport(2, summary, List.of(finding));
    TableHealthReport tableHealth = new TableHealthReport(1, summary, List.of(finding));

    assertEquals(
        2,
        new WorkbookReadResult.AutofilterHealthResult("autofilter-health", autofilterHealth)
            .analysis()
            .checkedAutofilterCount());
    assertEquals(
        1,
        new WorkbookReadResult.TableHealthResult("table-health", tableHealth)
            .analysis()
            .checkedTableCount());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.AutofilterHealthResult(" ", autofilterHealth));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.TableHealthResult("table-health", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AutofilterHealthReport(-1, summary, List.of(finding)));
    assertThrows(
        NullPointerException.class, () -> new TableHealthReport(1, null, List.of(finding)));
  }

  private static GridGrindResponse.CellStyleReport defaultStyle() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, java.math.BigDecimal.valueOf(11)),
            null,
            false,
            false),
        new CellFillReport(dev.erst.gridgrind.excel.ExcelFillPattern.NONE, null, null),
        new CellBorderReport(
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null)),
        new CellProtectionReport(true, false));
  }
}
