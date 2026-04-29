package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.json.JsonMapper;

/** Covers defaulting and convenience-ctor branches added across protocol DTOs. */
class ProtocolDefaultingCoverageTest {
  @Test
  void simpleCreatorsApplyRequestDefaultsWithoutAcceptingNullAtTheBoundary() {
    CalculationPolicyInput calculation = CalculationPolicyInput.create(null, null);
    ExecutionModeInput mode = ExecutionModeInput.create(null, null);
    ExecutionJournalInput journalInput = ExecutionJournalInput.create(null);
    FormulaEnvironmentInput environment = FormulaEnvironmentInput.create(null, null, null);
    SheetOutlineSummaryInput outlineSummary = SheetOutlineSummaryInput.create(null, null);
    SheetDisplayInput display = SheetDisplayInput.create(null, null, null, null, null);
    SheetDefaultsInput sheetDefaults = SheetDefaultsInput.create(null, null);
    HeaderFooterTextInput headerFooter = HeaderFooterTextInput.create(null, null, null);
    SheetPresentationInput presentation =
        SheetPresentationInput.create(null, null, null, null, null);
    WorkbookProtectionInput protection =
        WorkbookProtectionInput.create(null, null, null, null, null);
    AutofilterSortConditionInput sortCondition =
        AutofilterSortConditionInput.create("A1:A2", true, null, null, null);
    AutofilterFilterColumnInput filterColumn =
        AutofilterFilterColumnInput.create(
            1L, null, new AutofilterFilterCriterionInput.Values(List.of("Ready"), false));
    AutofilterSortStateInput sortState =
        AutofilterSortStateInput.create("A1:B2", null, null, null, List.of(sortCondition));
    FormulaUdfFunctionInput udfFunction =
        FormulaUdfFunctionInput.create("DOUBLE", 1, null, "ARG1*2");
    FormulaUdfFunctionInput explicitUdfFunction =
        FormulaUdfFunctionInput.create("TRIPLE", 1, 3, "ARG1*3");
    OoxmlSignatureInput signature =
        OoxmlSignatureInput.create("keys/signing.p12", "store-pass", null, null, null, null);
    OoxmlSignatureInput explicitSignature =
        OoxmlSignatureInput.create(
            "keys/signing.p12",
            "store-pass",
            "key-pass",
            "gridgrind-signing",
            ExcelOoxmlSignatureDigestAlgorithm.SHA512,
            "Signed workbook");
    RequestDoctorReport report =
        RequestDoctorReport.create(
            null,
            AnalysisSeverity.INFO,
            true,
            Optional.of(
                new RequestDoctorReport.Summary(
                    "NEW",
                    "NONE",
                    "FULL_XSSF",
                    "FULL_XSSF",
                    "DO_NOT_CALCULATE",
                    false,
                    false,
                    0,
                    0,
                    0,
                    0)),
            null,
            Optional.empty());

    assertEquals(new CalculationStrategyInput.DoNotCalculate(), calculation.strategy());
    assertFalse(calculation.markRecalculateOnOpen());
    assertTrue(mode.isDefault());
    assertEquals(ExecutionJournalLevel.NORMAL, journalInput.level());
    assertTrue(environment.isEmpty());
    assertEquals(SheetOutlineSummaryInput.defaults(), outlineSummary);
    assertEquals(SheetDisplayInput.defaults(), display);
    assertEquals(SheetDefaultsInput.defaults(), sheetDefaults);
    assertEquals(HeaderFooterTextInput.blank(), headerFooter);
    assertEquals(SheetPresentationInput.defaults(), presentation);
    assertFalse(protection.structureLocked());
    assertFalse(protection.windowsLocked());
    assertFalse(protection.revisionsLocked());
    assertEquals("", sortCondition.sortBy());
    assertTrue(filterColumn.showButton());
    assertFalse(sortState.caseSensitive());
    assertFalse(sortState.columnSort());
    assertEquals("", sortState.sortMethod());
    assertEquals(1, udfFunction.maximumArgumentCount());
    assertEquals(3, explicitUdfFunction.maximumArgumentCount());
    assertEquals("store-pass", signature.keyPassword());
    assertEquals(Optional.empty(), signature.alias());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signature.digestAlgorithm());
    assertEquals(Optional.empty(), signature.description());
    assertEquals("key-pass", explicitSignature.keyPassword());
    assertEquals(Optional.of("gridgrind-signing"), explicitSignature.alias());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA512, explicitSignature.digestAlgorithm());
    assertEquals(Optional.of("Signed workbook"), explicitSignature.description());
    assertEquals(GridGrindProtocolVersion.current(), report.protocolVersion());
    assertEquals(List.of(), report.warnings());
    assertEquals(Optional.empty(), ProtocolRgbColorSupport.normalizeRgbHex(null, null));
  }

  @Test
  void convenienceConstructorsCoverDefaultPlanAndReportBranches() {
    WorkbookPlan defaultPlan =
        new WorkbookPlan(
            GridGrindProtocolVersion.current(),
            "budget-defaults",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of());
    WorkbookPlan explicitExecutionPlan =
        new WorkbookPlan(
            GridGrindProtocolVersion.current(),
            "budget-execution",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicyInput(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            List.of());
    TableColumnReport tableColumn = new TableColumnReport(4L, "Owner");
    AutofilterSortConditionReport conditionReport =
        new AutofilterSortConditionReport("B2:B9", false, CellColorReport.rgb("#aabbcc"), null);
    AutofilterSortConditionReport createdConditionReport =
        AutofilterSortConditionReport.create("C2:C9", true, null, null, 2);
    AutofilterSortStateReport stateReport =
        new AutofilterSortStateReport("A1:F9", false, true, List.of(conditionReport));
    AutofilterSortStateReport createdStateReport =
        AutofilterSortStateReport.create(
            "A1:F9", true, false, null, List.of(createdConditionReport));
    DifferentialBorderSideInput borderSide =
        new DifferentialBorderSideInput(ExcelBorderStyle.THIN, Optional.empty());
    DifferentialBorderInput borderInput =
        new DifferentialBorderInput(
            Optional.of(borderSide),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    DifferentialBorderReport borderReport =
        new DifferentialBorderReport(
            Optional.empty(),
            Optional.of(new DifferentialBorderSideReport(ExcelBorderStyle.THIN, Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    AutofilterSortConditionInput explicitSortCondition =
        AutofilterSortConditionInput.create("D2:D9", false, "ICON", null, 1);
    AutofilterFilterColumnInput explicitFilterColumn =
        AutofilterFilterColumnInput.create(
            2L, false, new AutofilterFilterCriterionInput.Values(List.of("Queued"), true));
    AutofilterSortStateInput explicitSortState =
        AutofilterSortStateInput.create(
            "A1:D9", true, true, "PINYIN", List.of(explicitSortCondition));
    AutofilterSortConditionReport explicitConditionReport =
        AutofilterSortConditionReport.create("D2:D9", false, "ICON", null, 1);
    AutofilterSortStateReport explicitStateReport =
        AutofilterSortStateReport.create(
            "A1:D9", true, true, "PINYIN", List.of(explicitConditionReport));
    SheetDisplayInput explicitDisplay = SheetDisplayInput.create(false, true, false, true, true);
    SheetOutlineSummaryInput explicitOutline = SheetOutlineSummaryInput.create(false, true);
    SheetDefaultsInput explicitSheetDefaults = SheetDefaultsInput.create(18, 24.5d);
    SheetPresentationInput explicitPresentation =
        SheetPresentationInput.create(
            explicitDisplay,
            ColorInput.rgb("#aabbcc"),
            explicitOutline,
            explicitSheetDefaults,
            List.of(
                new IgnoredErrorInput(
                    "A1:A3", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))));
    FormulaEnvironmentInput explicitEnvironment =
        FormulaEnvironmentInput.create(
            List.of(new FormulaExternalWorkbookInput("Rates.xlsx", "tmp/rates.xlsx")),
            FormulaMissingWorkbookPolicy.USE_CACHED_VALUE,
            List.of());

    assertTrue(defaultPlan.execution().isDefault());
    assertTrue(defaultPlan.formulaEnvironment().isEmpty());
    assertEquals(
        ExecutionJournalLevel.VERBOSE, explicitExecutionPlan.execution().journal().level());
    assertTrue(explicitExecutionPlan.formulaEnvironment().isEmpty());
    assertEquals("", tableColumn.uniqueName());
    assertEquals("", tableColumn.totalsRowLabel());
    assertEquals("", conditionReport.sortBy());
    assertEquals("", createdConditionReport.sortBy());
    assertEquals("", stateReport.sortMethod());
    assertEquals("", createdStateReport.sortMethod());
    assertEquals(borderSide, borderInput.all().orElseThrow());
    assertEquals(ExcelBorderStyle.THIN, borderReport.top().orElseThrow().style());
    assertEquals("ICON", explicitSortCondition.sortBy());
    assertFalse(explicitFilterColumn.showButton());
    assertEquals("PINYIN", explicitSortState.sortMethod());
    assertEquals("ICON", explicitConditionReport.sortBy());
    assertEquals("PINYIN", explicitStateReport.sortMethod());
    assertEquals(explicitDisplay, explicitPresentation.display());
    assertEquals(explicitOutline, explicitPresentation.outlineSummary());
    assertEquals(explicitSheetDefaults, explicitPresentation.sheetDefaults());
    assertEquals(1, explicitPresentation.ignoredErrors().size());
    assertEquals(
        FormulaMissingWorkbookPolicy.USE_CACHED_VALUE, explicitEnvironment.missingWorkbookPolicy());
  }

  @Test
  void chartCreatorsAndConvenienceConstructorsSupplyExplicitDefaults() {
    ChartInput.Series explicitSeries =
        ChartInput.Series.create(
            null,
            new ChartInput.DataSource.StringLiteral(List.of("Jan", "Feb")),
            new ChartInput.DataSource.NumericLiteral(List.of(10.0d, 18.0d)),
            true,
            ExcelChartMarkerStyle.DIAMOND,
            (short) 8,
            12L);
    ChartInput.Series convenienceSeries =
        new ChartInput.Series(
            new ChartInput.DataSource.Reference("Categories"),
            new ChartInput.DataSource.Reference("Values"),
            false,
            ExcelChartMarkerStyle.CIRCLE,
            (short) 6,
            null);
    ChartInput.Axis defaultAxis =
        new ChartInput.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO);
    List<ChartInput.Axis> explicitAxes =
        List.of(
            defaultAxis,
            new ChartInput.Axis(
                ExcelChartAxisKind.VALUE,
                ExcelChartAxisPosition.LEFT,
                ExcelChartAxisCrosses.AUTO_ZERO,
                false));
    ChartInput.Area explicitArea =
        ChartInput.Area.create(
            true, ExcelChartGrouping.PERCENT_STACKED, explicitAxes, List.of(explicitSeries));
    ChartInput.Area convenienceArea =
        new ChartInput.Area(false, ExcelChartGrouping.STANDARD, List.of(explicitSeries));
    ChartInput.Area3D explicitArea3D =
        ChartInput.Area3D.create(
            false, ExcelChartGrouping.STACKED, 24, explicitAxes, List.of(explicitSeries));
    ChartInput.Area3D convenienceArea3D =
        new ChartInput.Area3D(false, ExcelChartGrouping.STANDARD, 16, List.of(explicitSeries));
    ChartInput.Bar3D explicitBar3D =
        ChartInput.Bar3D.create(
            false,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.PERCENT_STACKED,
            12,
            64,
            ExcelChartBarShape.CONE,
            explicitAxes,
            List.of(explicitSeries));
    ChartInput.Bar3D convenienceBar3D =
        new ChartInput.Bar3D(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            18,
            90,
            ExcelChartBarShape.BOX,
            List.of(explicitSeries));
    ChartInput.Line explicitLine =
        ChartInput.Line.create(
            true, ExcelChartGrouping.STACKED, explicitAxes, List.of(explicitSeries));
    ChartInput.Line convenienceLine =
        new ChartInput.Line(false, ExcelChartGrouping.STANDARD, List.of(explicitSeries));
    ChartInput.Line3D explicitLine3D =
        ChartInput.Line3D.create(
            true, ExcelChartGrouping.PERCENT_STACKED, 32, explicitAxes, List.of(explicitSeries));
    ChartInput.Line3D convenienceLine3D =
        new ChartInput.Line3D(false, ExcelChartGrouping.STANDARD, 8, List.of(explicitSeries));
    ChartInput.Radar explicitRadar =
        ChartInput.Radar.create(
            true, ExcelChartRadarStyle.MARKER, explicitAxes, List.of(explicitSeries));
    ChartInput.Radar convenienceRadar =
        new ChartInput.Radar(false, ExcelChartRadarStyle.STANDARD, List.of(explicitSeries));
    ChartInput.Scatter explicitScatter =
        ChartInput.Scatter.create(
            true, ExcelChartScatterStyle.SMOOTH, explicitAxes, List.of(explicitSeries));
    ChartInput.Scatter convenienceScatter =
        new ChartInput.Scatter(false, ExcelChartScatterStyle.LINE, List.of(explicitSeries));
    ChartInput.Surface explicitSurface =
        ChartInput.Surface.create(true, true, explicitAxes, List.of(explicitSeries));
    ChartInput.Surface convenienceSurface =
        new ChartInput.Surface(false, false, List.of(explicitSeries));
    ChartInput.Surface3D explicitSurface3D =
        ChartInput.Surface3D.create(true, true, explicitAxes, List.of(explicitSeries));
    ChartInput.Surface3D convenienceSurface3D =
        new ChartInput.Surface3D(false, false, List.of(explicitSeries));
    ChartInput createdChart =
        ChartInput.create(
            "CreatedChart", twoCellAnchor(), null, null, null, null, List.of(explicitArea));
    ChartInput convenienceChart =
        new ChartInput("ConvenienceChart", twoCellAnchor(), List.of(explicitLine));

    assertTrue(defaultAxis.visible());
    assertTrue(convenienceSeries.title() instanceof ChartInput.Title.None);
    assertEquals(2, explicitArea.axes().size());
    assertEquals(2, convenienceArea.axes().size());
    assertEquals(2, explicitArea3D.axes().size());
    assertEquals(2, convenienceArea3D.axes().size());
    assertEquals(2, explicitBar3D.axes().size());
    assertEquals(2, convenienceBar3D.axes().size());
    assertEquals(2, explicitLine.axes().size());
    assertEquals(2, convenienceLine.axes().size());
    assertEquals(2, explicitLine3D.axes().size());
    assertEquals(2, convenienceLine3D.axes().size());
    assertEquals(2, explicitRadar.axes().size());
    assertEquals(2, convenienceRadar.axes().size());
    assertEquals(2, explicitScatter.axes().size());
    assertEquals(2, convenienceScatter.axes().size());
    assertEquals(2, explicitSurface.axes().size());
    assertEquals(3, convenienceSurface.axes().size());
    assertEquals(2, explicitSurface3D.axes().size());
    assertEquals(3, convenienceSurface3D.axes().size());
    assertTrue(createdChart.title() instanceof ChartInput.Title.None);
    assertEquals(
        new ChartInput.Legend.Visible(ExcelChartLegendPosition.RIGHT), createdChart.legend());
    assertEquals(ExcelChartDisplayBlanksAs.GAP, createdChart.displayBlanksAs());
    assertTrue(createdChart.plotOnlyVisibleCells());
    assertTrue(convenienceChart.title() instanceof ChartInput.Title.None);
  }

  @Test
  void delegatingCreatorsCoverPrivateDefaultingShapes() throws IOException {
    JsonMapper mapper = JsonMapper.builder().build();
    PrintSetupInput printSetup =
        mapper.readValue(
            """
            {
              "margins": null,
              "printGridlines": null,
              "horizontallyCentered": null,
              "verticallyCentered": null,
              "paperSize": null,
              "draft": null,
              "blackAndWhite": null,
              "copies": null,
              "useFirstPageNumber": null,
              "firstPageNumber": null,
              "rowBreaks": null,
              "columnBreaks": null
            }
            """,
            PrintSetupInput.class);
    PrintLayoutInput printLayout =
        mapper.readValue(
            """
            {
              "setup": {
                "margins": {
                  "left": 0.6,
                  "right": 0.6,
                  "top": 0.7,
                  "bottom": 0.7,
                  "header": 0.3,
                  "footer": 0.3
                },
                "printGridlines": true,
                "horizontallyCentered": false,
                "verticallyCentered": false,
                "paperSize": 9,
                "draft": false,
                "blackAndWhite": false,
                "copies": 2,
                "useFirstPageNumber": false,
                "firstPageNumber": 0,
                "rowBreaks": [],
                "columnBreaks": []
              }
            }
            """,
            PrintLayoutInput.class);
    PrintLayoutInput conveniencePrintLayout =
        new PrintLayoutInput(
            new PrintAreaInput.None(),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingInput.Automatic(),
            new PrintTitleRowsInput.None(),
            new PrintTitleColumnsInput.None(),
            HeaderFooterTextInput.blank(),
            HeaderFooterTextInput.blank());
    PrintLayoutInput createdPrintLayout =
        PrintLayoutInput.create(
            new PrintAreaInput.None(),
            ExcelPrintOrientation.LANDSCAPE,
            new PrintScalingInput.Automatic(),
            new PrintTitleRowsInput.None(),
            new PrintTitleColumnsInput.None(),
            HeaderFooterTextInput.blank(),
            HeaderFooterTextInput.blank());
    SignatureLineInput signatureLine =
        mapper.readValue(
            """
            {
              "name": "OpsSignature",
              "anchor": {
                "type": "TWO_CELL",
                "from": { "columnIndex": 1, "rowIndex": 0, "dx": 1, "dy": 0 },
                "to": { "columnIndex": 5, "rowIndex": 10, "dx": 1, "dy": 0 },
                "behavior": "MOVE_AND_RESIZE"
              },
              "allowComments": null,
              "suggestedSigner": "Ada Lovelace"
            }
            """,
            SignatureLineInput.class);
    TableEntryReport tableReport =
        mapper.readValue(
            """
            {
              "name": "BudgetTable",
              "sheetName": "Budget",
              "range": "A1:B4",
              "headerRowCount": 1,
              "totalsRowCount": 0,
              "columnNames": ["Owner"],
              "columns": [{ "id": 0, "name": "Owner" }],
              "style": { "type": "NONE" },
              "hasAutofilter": true,
              "comment": null,
              "published": false,
              "insertRow": false,
              "insertRowShift": false,
              "headerRowCellStyle": null,
              "dataCellStyle": null,
              "totalsRowCellStyle": null
            }
            """,
            TableEntryReport.class);
    ExecutionJournal journal =
        mapper.readValue(
            """
            {
              "planId": "journal-plan",
              "level": null,
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "validation": { "status": "NOT_STARTED", "durationMillis": 0 },
              "inputResolution": { "status": "NOT_STARTED", "durationMillis": 0 },
              "open": { "status": "NOT_STARTED", "durationMillis": 0 },
              "calculation": {
                "preflight": { "status": "NOT_STARTED", "durationMillis": 0 },
                "execution": { "status": "NOT_STARTED", "durationMillis": 0 }
              },
              "persistencePhase": { "status": "NOT_STARTED", "durationMillis": 0 },
              "close": { "status": "NOT_STARTED", "durationMillis": 0 },
              "steps": [],
              "warnings": null,
              "outcome": {
                "status": "SUCCEEDED",
                "plannedStepCount": 0,
                "completedStepCount": 0,
                "durationMillis": 0
              },
              "events": null
            }
            """,
            ExecutionJournal.class);

    assertEquals(PrintSetupInput.defaults(), printSetup);
    assertEquals(ExcelPrintOrientation.PORTRAIT, printLayout.orientation());
    assertTrue(printLayout.setup().printGridlines());
    assertEquals(PrintSetupInput.defaults(), conveniencePrintLayout.setup());
    assertEquals(PrintSetupInput.defaults(), createdPrintLayout.setup());
    assertTrue(signatureLine.allowComments());
    assertEquals("", tableReport.comment());
    assertEquals("", tableReport.headerRowCellStyle());
    assertEquals("", tableReport.dataCellStyle());
    assertEquals("", tableReport.totalsRowCellStyle());
    assertEquals(ExecutionJournalLevel.NORMAL, journal.level());
    assertEquals(List.of(), journal.warnings());
    assertEquals(List.of(), journal.events());
  }

  private static DrawingAnchorInput.TwoCell twoCellAnchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(1, 0, 1, 0),
        new DrawingMarkerInput(5, 0, 10, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }
}
