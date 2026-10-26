package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartMarkerStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartRadarStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelChartScatterStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.DrawingGeometryFixtures;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.function.Function;
import org.junit.jupiter.api.Test;

/** Covers the remaining XLSX conversion and source-resolution branches end to end. */
class SpreadsheetSurfaceConverterCoverageTest {
  @Test
  void workbookCommandConverterCoversEveryExtendedChartPlotAndNullSignaturePreview() {
    ChartSpec chart =
        new ChartSpec(
            "ExtendedChart",
            inputAnchor(),
            new ChartTitleSpec.Text(TextSourceSpec.inline("Extended")),
            new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP_RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            inputPlots(inlineSeries(TextSourceSpec.inline("Series"))));
    ResolvedChart definition = WorkbookCommandDrawingInputConverter.toResolvedChart(chart);
    WorkbookDrawingCommand.SetChart chartCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName(new SheetName("Ops")),
                new DrawingMutationAction.SetChart(chart)));
    ResolvedSignatureLine signatureLine =
        WorkbookCommandDrawingInputConverter.toResolvedSignatureLine(
            new SignatureLineSpec(
                "OpsSignature",
                inputAnchor(),
                false,
                java.util.Optional.of("Review before signing."),
                java.util.Optional.of("Ada Lovelace"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.of("invalid"),
                java.util.Optional.empty()));

    assertEquals(expectedDefinitionPlotTypes(), plotTypes(definition.plots()));
    assertEquals(
        List.of("Jan", "Feb"),
        assertInstanceOf(
                ChartDataSourceSpec.StringLiteral.class,
                assertInstanceOf(ResolvedChart.Area.class, definition.plots().getFirst())
                    .series()
                    .getFirst()
                    .categories())
            .values());
    assertEquals(
        List.of(10.0d, 18.0d),
        assertInstanceOf(
                ChartDataSourceSpec.NumericLiteral.class,
                assertInstanceOf(ResolvedChart.Area.class, definition.plots().getFirst())
                    .series()
                    .getFirst()
                    .values())
            .values());
    assertEquals(
        Optional.of(ExcelChartBarShape.CONE),
        assertInstanceOf(ResolvedChart.Bar3D.class, definition.plots().get(3)).shape());
    assertEquals(chartCommand.chart(), definition);
    assertEquals(Optional.empty(), signatureLine.plainSignatureFormat());
    assertEquals(Optional.empty(), signatureLine.plainSignature());
  }

  @Test
  void inspectionResultConverterCoversEveryExtendedChartPlotAndCustomXmlLinkedTables() {
    CustomXmlMappingReport mappingSnapshot =
        new CustomXmlMappingReport(
            1L,
            "CORSO_mapping",
            "CORSO",
            "Schema1",
            new CustomXmlMappingReport.Settings(false, true, false, true, true),
            new CustomXmlMappingReport.Schema(
                "urn:gridgrind:test", "XSD", "schema.xsd", "<xsd:schema/>"),
            new CustomXmlDataBindingReport("binding", true, 5L, "binding.xml", 2L),
            List.of(
                new CustomXmlLinkedCellReport(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Foglio1"),
                    new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                    "/CORSO/NOME",
                    "string")),
            List.of(
                new CustomXmlLinkedTableReport(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Foglio1"),
                    "Table1",
                    "CourseTable",
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B4"),
                    "/CORSO/RIGHE/RIGA")));
    ChartReport snapshot =
        new ChartReport(
            "ExtendedChart",
            DrawingGeometryFixtures.report(excelAnchor()),
            new ChartReport.Title.Text("Extended"),
            new ChartReport.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            snapshotPlots(snapshotSeries(new ChartReport.Title.Text("Series"))));
    ChartReport chartReport =
        assertInstanceOf(
                dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.ChartsResult.class,
                InspectionResultConverter.toReadResult(
                    new dev.erst.gridgrind.excel.WorkbookDrawingResult.ChartsResult(
                        new StepId("charts"), "Foglio1", List.of(snapshot))))
            .charts()
            .getFirst();
    WorkbookInspectionResult.CustomXmlMappingsResult mappings =
        assertInstanceOf(
            WorkbookInspectionResult.CustomXmlMappingsResult.class,
            InspectionResultConverter.toReadResult(
                new WorkbookCoreResult.CustomXmlMappingsResult(
                    new StepId("custom-xml"), List.of(mappingSnapshot))));

    assertEquals(expectedReportPlotTypes(), plotTypes(chartReport.plots()));
    assertEquals(
        List.of("Jan", "Feb"),
        assertInstanceOf(
                ChartReport.DataSource.StringLiteral.class,
                assertInstanceOf(ChartReport.Area.class, chartReport.plots().getFirst())
                    .series()
                    .getFirst()
                    .categories())
            .values());
    assertEquals(
        "CourseTable", mappings.mappings().getFirst().linkedTables().getFirst().tableDisplayName());
    assertEquals(
        "/CORSO/RIGHE/RIGA",
        mappings.mappings().getFirst().linkedTables().getFirst().commonXPath());
  }

  @Test
  void sourceBackedHelpersCoverExtendedChartPlotFamiliesAndStableNullBranches() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-chart-surface-");
    Files.writeString(
        workingDirectory.resolve("series-title.txt"), "Resolved Series", StandardCharsets.UTF_8);

    ChartSpec fileBackedChart =
        new ChartSpec(
            "ExtendedChart",
            inputAnchor(),
            new ChartTitleSpec.Text(TextSourceSpec.inline("Extended")),
            new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            inputPlots(inlineSeries(TextSourceSpec.utf8File("series-title.txt"))));
    ChartSpec standardInputChart =
        new ChartSpec(
            "ExtendedChart",
            inputAnchor(),
            new ChartTitleSpec.None(),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            inputPlots(inlineSeries(TextSourceSpec.standardInput())));
    DrawingMutationAction.SetChart chartAction =
        new DrawingMutationAction.SetChart(fileBackedChart);
    DrawingMutationAction.SetSignatureLine signatureAction =
        new DrawingMutationAction.SetSignatureLine(
            new SignatureLineSpec(
                "OpsSignature",
                inputAnchor(),
                true,
                java.util.Optional.empty(),
                java.util.Optional.of("Ada Lovelace"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty()));
    DrawingMutationAction.SetChart stableChartAction =
        new DrawingMutationAction.SetChart(
            new ChartSpec(
                "StableChart",
                inputAnchor(),
                new ChartTitleSpec.None(),
                new ChartLegendSpec.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                inputPlots(inlineSeries(TextSourceSpec.inline("Stable Series")))));
    DrawingMutationAction.SetSignatureLine stableSignatureAction =
        new DrawingMutationAction.SetSignatureLine(signatureLineWithInlineBinary());
    StructuredMutationAction.ImportCustomXmlMapping inlineImportAction =
        new StructuredMutationAction.ImportCustomXmlMapping(
            new CustomXmlImportSpec(
                new CustomXmlMappingLocator(
                    java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                TextSourceSpec.inline("<root/>")));

    for (ChartPlotSpec plot : standardInputChart.plots()) {
      assertTrue(SourceBackedInputRequirements.requiresStandardInput(plot));
    }
    assertTrue(SourceBackedInputRequirements.requiresStandardInput(standardInputChart));
    assertTrue(SourceBackedInputRequirements.requiresStandardInput(signatureLineWithStdIn()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(signatureLineWithInlineBinary()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new SignatureLineSpec(
                "OpsSignature",
                inputAnchor(),
                true,
                java.util.Optional.empty(),
                java.util.Optional.of("Ada Lovelace"),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                java.util.Optional.empty())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new CustomXmlImportSpec(
                new CustomXmlMappingLocator(
                    java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                TextSourceSpec.standardInput())));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new CustomXmlImportSpec(
                new CustomXmlMappingLocator(
                    java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                TextSourceSpec.inline("<root/>"))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetSignatureLine(signatureLineWithStdIn())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.ImportCustomXmlMapping(
                new CustomXmlImportSpec(
                    new CustomXmlMappingLocator(
                        java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping")),
                    TextSourceSpec.standardInput()))));

    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("set-chart"),
                    new SheetSelector.ByName(new SheetName("Ops")),
                    chartAction),
                new MutationStep(
                    new StepId("set-signature"),
                    new SheetSelector.ByName(new SheetName("Ops")),
                    signatureAction),
                new MutationStep(
                    new StepId("import-custom-xml"),
                    new WorkbookSelector.Current(),
                    inlineImportAction)));
    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            plan,
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory, plan, "stdin".getBytes(StandardCharsets.UTF_8)));

    DrawingMutationAction.SetChart resolvedChartAction =
        assertInstanceOf(
            DrawingMutationAction.SetChart.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(0)).action());
    assertEquals(expectedInputPlotTypes(), plotTypes(resolvedChartAction.chart().plots()));
    for (ChartPlotSpec plot : resolvedChartAction.chart().plots()) {
      ChartSeriesSpec series = firstSeries(plot);
      ChartTitleSpec.Text title = assertInstanceOf(ChartTitleSpec.Text.class, series.title());
      assertEquals(
          "Resolved Series", assertInstanceOf(TextSourceSpec.Inline.class, title.source()).text());
    }
    assertSame(
        signatureAction, assertInstanceOf(MutationStep.class, resolved.steps().get(1)).action());
    assertSame(
        inlineImportAction, assertInstanceOf(MutationStep.class, resolved.steps().get(2)).action());

    WorkbookPlan stablePlan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.core.plan.ExecutionPolicySpec.defaults(),
            dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("stable-chart"),
                    new SheetSelector.ByName(new SheetName("Ops")),
                    stableChartAction),
                new MutationStep(
                    new StepId("stable-signature"),
                    new SheetSelector.ByName(new SheetName("Ops")),
                    stableSignatureAction)));
    WorkbookPlan stableResolved =
        SourceBackedPlanResolver.resolve(
            stablePlan,
            ExecutionInputBindingsFixtureSupport.bindings(
                workingDirectory, stablePlan, "stdin".getBytes(StandardCharsets.UTF_8)));

    assertSame(
        stableChartAction,
        assertInstanceOf(MutationStep.class, stableResolved.steps().get(0)).action());
    assertSame(
        stableSignatureAction,
        assertInstanceOf(MutationStep.class, stableResolved.steps().get(1)).action());
  }

  private static SignatureLineSpec signatureLineWithStdIn() {
    return new SignatureLineSpec(
        "OpsSignature",
        inputAnchor(),
        true,
        java.util.Optional.empty(),
        java.util.Optional.of("Ada Lovelace"),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.of(
            new PictureDataSpec(ExcelPictureFormat.PNG, BinarySourceSpec.standardInput())));
  }

  private static SignatureLineSpec signatureLineWithInlineBinary() {
    return new SignatureLineSpec(
        "OpsSignature",
        inputAnchor(),
        true,
        java.util.Optional.empty(),
        java.util.Optional.of("Ada Lovelace"),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.of(
            new PictureDataSpec(
                ExcelPictureFormat.PNG, BinarySourceSpec.inlineBase64("cGF5bG9hZA=="))));
  }

  private static DrawingAnchorSpec.TwoCell inputAnchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(1, 2, 0, 0),
        new DrawingMarkerSpec(8, 14, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static DrawingAnchorSpec.TwoCell excelAnchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(1, 2, 0, 0),
        new DrawingMarkerSpec(8, 14, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static ChartSeriesSpec inlineSeries(TextSourceSpec title) {
    return new ChartSeriesSpec(
        new ChartTitleSpec.Text(title),
        new ChartDataSourceSpec.StringLiteral(List.of("Jan", "Feb")),
        new ChartDataSourceSpec.NumericLiteral(List.of(10.0d, 18.0d)),
        Optional.of(true),
        Optional.of(ExcelChartMarkerStyle.DIAMOND),
        Optional.of((short) 6),
        Optional.of(4L));
  }

  private static ChartReport.Series snapshotSeries(ChartReport.Title title) {
    return new ChartReport.Series(
        title,
        new ChartReport.DataSource.StringLiteral(List.of("Jan", "Feb")),
        new ChartReport.DataSource.NumericLiteral(Optional.of("0.0"), List.of("10", "18")),
        Optional.of(true),
        Optional.of(ExcelChartMarkerStyle.DIAMOND),
        Optional.of((short) 6),
        Optional.of(4L));
  }

  private static List<ChartPlotSpec> inputPlots(ChartSeriesSpec series) {
    return plotVariants().stream().map(variant -> variant.inputFactory().apply(series)).toList();
  }

  private static List<ChartReport.Plot> snapshotPlots(ChartReport.Series series) {
    return plotVariants().stream().map(variant -> variant.snapshotFactory().apply(series)).toList();
  }

  private static List<PlotVariant> plotVariants() {
    return List.of(
        new PlotVariant(
            series ->
                new ChartPlotSpec.Area(
                    false, ExcelChartGrouping.STANDARD, inputAxes(), List.of(series)),
            series ->
                new ChartReport.Area(
                    false, ExcelChartGrouping.STANDARD, snapshotAxes(), List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Area3D(
                    false,
                    ExcelChartGrouping.PERCENT_STACKED,
                    Optional.of(24),
                    inputAxes(),
                    List.of(series)),
            series ->
                new ChartReport.Area3D(
                    false,
                    ExcelChartGrouping.PERCENT_STACKED,
                    Optional.of(24),
                    snapshotAxes(),
                    List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Bar(
                    false,
                    ExcelChartBarDirection.COLUMN,
                    ExcelChartBarGrouping.CLUSTERED,
                    Optional.of(60),
                    Optional.of(0),
                    inputAxes(),
                    List.of(series)),
            series ->
                new ChartReport.Bar(
                    false,
                    ExcelChartBarDirection.COLUMN,
                    ExcelChartBarGrouping.CLUSTERED,
                    Optional.of(60),
                    Optional.of(0),
                    snapshotAxes(),
                    List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Bar3D(
                    true,
                    ExcelChartBarDirection.BAR,
                    ExcelChartBarGrouping.STACKED,
                    Optional.of(32),
                    Optional.of(88),
                    Optional.of(ExcelChartBarShape.CONE),
                    inputAxes(),
                    List.of(series)),
            series ->
                new ChartReport.Bar3D(
                    true,
                    ExcelChartBarDirection.BAR,
                    ExcelChartBarGrouping.STACKED,
                    Optional.of(32),
                    Optional.of(88),
                    Optional.of(ExcelChartBarShape.CONE),
                    snapshotAxes(),
                    List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Doughnut(true, Optional.of(30), Optional.of(55), List.of(series)),
            series ->
                new ChartReport.Doughnut(true, Optional.of(30), Optional.of(55), List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Line(
                    false, ExcelChartGrouping.STANDARD, inputAxes(), List.of(series)),
            series ->
                new ChartReport.Line(
                    false, ExcelChartGrouping.STANDARD, snapshotAxes(), List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Line3D(
                    false,
                    ExcelChartGrouping.STANDARD,
                    Optional.of(18),
                    inputAxes(),
                    List.of(series)),
            series ->
                new ChartReport.Line3D(
                    false,
                    ExcelChartGrouping.STANDARD,
                    Optional.of(18),
                    snapshotAxes(),
                    List.of(series))),
        new PlotVariant(
            series -> new ChartPlotSpec.Pie(true, Optional.of(90), List.of(series)),
            series -> new ChartReport.Pie(true, Optional.of(90), List.of(series))),
        new PlotVariant(
            series -> new ChartPlotSpec.Pie3D(true, List.of(series)),
            series -> new ChartReport.Pie3D(true, List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Radar(
                    false, ExcelChartRadarStyle.FILLED, inputAxes(), List.of(series)),
            series ->
                new ChartReport.Radar(
                    false, ExcelChartRadarStyle.FILLED, snapshotAxes(), List.of(series))),
        new PlotVariant(
            series ->
                new ChartPlotSpec.Scatter(
                    false,
                    ExcelChartScatterStyle.SMOOTH_MARKER,
                    scatterInputAxes(),
                    List.of(series)),
            series ->
                new ChartReport.Scatter(
                    false,
                    ExcelChartScatterStyle.SMOOTH_MARKER,
                    scatterSnapshotAxes(),
                    List.of(series))),
        new PlotVariant(
            series -> new ChartPlotSpec.Surface(false, true, surfaceInputAxes(), List.of(series)),
            series -> new ChartReport.Surface(false, true, surfaceSnapshotAxes(), List.of(series))),
        new PlotVariant(
            series -> new ChartPlotSpec.Surface3D(true, false, surfaceInputAxes(), List.of(series)),
            series ->
                new ChartReport.Surface3D(true, false, surfaceSnapshotAxes(), List.of(series))));
  }

  private static List<ChartAxisSpec> inputAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisSpec> scatterInputAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisSpec> surfaceInputAxes() {
    return List.of(
        new ChartAxisSpec(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisSpec(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartReport.Axis> snapshotAxes() {
    return List.of(
        new ChartReport.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartReport.Axis> scatterSnapshotAxes() {
    return List.of(
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartReport.Axis> surfaceSnapshotAxes() {
    return List.of(
        new ChartReport.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartReport.Axis(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<Class<?>> expectedDefinitionPlotTypes() {
    return List.of(
        ResolvedChart.Area.class,
        ResolvedChart.Area3D.class,
        ResolvedChart.Bar.class,
        ResolvedChart.Bar3D.class,
        ResolvedChart.Doughnut.class,
        ResolvedChart.Line.class,
        ResolvedChart.Line3D.class,
        ResolvedChart.Pie.class,
        ResolvedChart.Pie3D.class,
        ResolvedChart.Radar.class,
        ResolvedChart.Scatter.class,
        ResolvedChart.Surface.class,
        ResolvedChart.Surface3D.class);
  }

  private static List<Class<?>> expectedReportPlotTypes() {
    return List.of(
        ChartReport.Area.class,
        ChartReport.Area3D.class,
        ChartReport.Bar.class,
        ChartReport.Bar3D.class,
        ChartReport.Doughnut.class,
        ChartReport.Line.class,
        ChartReport.Line3D.class,
        ChartReport.Pie.class,
        ChartReport.Pie3D.class,
        ChartReport.Radar.class,
        ChartReport.Scatter.class,
        ChartReport.Surface.class,
        ChartReport.Surface3D.class);
  }

  private static List<Class<?>> expectedInputPlotTypes() {
    return List.of(
        ChartPlotSpec.Area.class,
        ChartPlotSpec.Area3D.class,
        ChartPlotSpec.Bar.class,
        ChartPlotSpec.Bar3D.class,
        ChartPlotSpec.Doughnut.class,
        ChartPlotSpec.Line.class,
        ChartPlotSpec.Line3D.class,
        ChartPlotSpec.Pie.class,
        ChartPlotSpec.Pie3D.class,
        ChartPlotSpec.Radar.class,
        ChartPlotSpec.Scatter.class,
        ChartPlotSpec.Surface.class,
        ChartPlotSpec.Surface3D.class);
  }

  private record PlotVariant(
      Function<ChartSeriesSpec, ChartPlotSpec> inputFactory,
      Function<ChartReport.Series, ChartReport.Plot> snapshotFactory) {}

  private static List<Class<?>> plotTypes(List<?> plots) {
    return plots.stream().<Class<?>>map(Object::getClass).toList();
  }

  private static ChartSeriesSpec firstSeries(ChartPlotSpec plot) {
    return switch (plot) {
      case ChartPlotSpec.Area area -> area.series().getFirst();
      case ChartPlotSpec.Area3D area3D -> area3D.series().getFirst();
      case ChartPlotSpec.Bar bar -> bar.series().getFirst();
      case ChartPlotSpec.Bar3D bar3D -> bar3D.series().getFirst();
      case ChartPlotSpec.Doughnut doughnut -> doughnut.series().getFirst();
      case ChartPlotSpec.Line line -> line.series().getFirst();
      case ChartPlotSpec.Line3D line3D -> line3D.series().getFirst();
      case ChartPlotSpec.Pie pie -> pie.series().getFirst();
      case ChartPlotSpec.Pie3D pie3D -> pie3D.series().getFirst();
      case ChartPlotSpec.Radar radar -> radar.series().getFirst();
      case ChartPlotSpec.Scatter scatter -> scatter.series().getFirst();
      case ChartPlotSpec.Surface surface -> surface.series().getFirst();
      case ChartPlotSpec.Surface3D surface3D -> surface3D.series().getFirst();
    };
  }
}
