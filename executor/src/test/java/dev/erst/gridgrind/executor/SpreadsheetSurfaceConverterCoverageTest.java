package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.ChartAxisInput;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlDataBindingSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlLinkedCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlLinkedTableSnapshot;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;
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
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers the remaining XLSX conversion and source-resolution branches end to end. */
class SpreadsheetSurfaceConverterCoverageTest {
  @Test
  void workbookCommandConverterCoversEveryExtendedChartPlotAndNullSignaturePreview() {
    ChartInput chart =
        new ChartInput(
            "ExtendedChart",
            inputAnchor(),
            new ChartTitleInput.Text(TextSourceInput.inline("Extended")),
            new ChartLegendInput.Visible(ExcelChartLegendPosition.TOP_RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            inputPlots(inlineSeries(TextSourceInput.inline("Series"))));
    ExcelChartDefinition definition = WorkbookCommandConverter.toExcelChartDefinition(chart);
    WorkbookDrawingCommand.SetChart chartCommand =
        assertInstanceOf(
            WorkbookDrawingCommand.SetChart.class,
            WorkbookCommandConverter.toCommand(
                new SheetSelector.ByName("Ops"), new DrawingMutationAction.SetChart(chart)));
    ExcelSignatureLineDefinition signatureLine =
        WorkbookCommandConverter.toExcelSignatureLineDefinition(
            new SignatureLineInput(
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
                ExcelChartDefinition.DataSource.StringLiteral.class,
                assertInstanceOf(ExcelChartDefinition.Area.class, definition.plots().getFirst())
                    .series()
                    .getFirst()
                    .categories())
            .values());
    assertEquals(
        List.of(10.0d, 18.0d),
        assertInstanceOf(
                ExcelChartDefinition.DataSource.NumericLiteral.class,
                assertInstanceOf(ExcelChartDefinition.Area.class, definition.plots().getFirst())
                    .series()
                    .getFirst()
                    .values())
            .values());
    assertEquals(
        ExcelChartBarShape.CONE,
        assertInstanceOf(ExcelChartDefinition.Bar3D.class, definition.plots().get(3)).shape());
    assertEquals(chartCommand.chart(), definition);
    assertNull(signatureLine.plainSignatureFormat());
    assertNull(signatureLine.plainSignature());
  }

  @Test
  void inspectionResultConverterCoversEveryExtendedChartPlotAndCustomXmlLinkedTables() {
    ExcelCustomXmlMappingSnapshot mappingSnapshot =
        new ExcelCustomXmlMappingSnapshot(
            1L,
            "CORSO_mapping",
            "CORSO",
            "Schema1",
            false,
            true,
            false,
            true,
            true,
            "urn:gridgrind:test",
            "XSD",
            "schema.xsd",
            "<xsd:schema/>",
            new ExcelCustomXmlDataBindingSnapshot("binding", true, 5L, "binding.xml", 2L),
            List.of(new ExcelCustomXmlLinkedCellSnapshot("Foglio1", "A1", "/CORSO/NOME", "string")),
            List.of(
                new ExcelCustomXmlLinkedTableSnapshot(
                    "Foglio1", "Table1", "CourseTable", "A1:B4", "/CORSO/RIGHE/RIGA")));
    ExcelChartSnapshot snapshot =
        new ExcelChartSnapshot(
            "ExtendedChart",
            excelAnchor(),
            new ExcelChartSnapshot.Title.Text("Extended"),
            new ExcelChartSnapshot.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            snapshotPlots(snapshotSeries(new ExcelChartSnapshot.Title.Text("Series"))));
    ChartReport chartReport = InspectionResultDrawingReportSupport.toChartReport(snapshot);
    InspectionResult.CustomXmlMappingsResult mappings =
        assertInstanceOf(
            InspectionResult.CustomXmlMappingsResult.class,
            InspectionResultConverter.toReadResult(
                new WorkbookCoreResult.CustomXmlMappingsResult(
                    "custom-xml", List.of(mappingSnapshot))));

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

    ChartInput fileBackedChart =
        new ChartInput(
            "ExtendedChart",
            inputAnchor(),
            new ChartTitleInput.Text(TextSourceInput.inline("Extended")),
            new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            inputPlots(inlineSeries(TextSourceInput.utf8File("series-title.txt"))));
    ChartInput standardInputChart =
        new ChartInput(
            "ExtendedChart",
            inputAnchor(),
            new ChartTitleInput.None(),
            new ChartLegendInput.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            inputPlots(inlineSeries(TextSourceInput.standardInput())));
    DrawingMutationAction.SetChart chartAction =
        new DrawingMutationAction.SetChart(fileBackedChart);
    DrawingMutationAction.SetSignatureLine signatureAction =
        new DrawingMutationAction.SetSignatureLine(
            new SignatureLineInput(
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
            new ChartInput(
                "StableChart",
                inputAnchor(),
                new ChartTitleInput.None(),
                new ChartLegendInput.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                inputPlots(inlineSeries(TextSourceInput.inline("Stable Series")))));
    DrawingMutationAction.SetSignatureLine stableSignatureAction =
        new DrawingMutationAction.SetSignatureLine(signatureLineWithInlineBinary());
    StructuredMutationAction.ImportCustomXmlMapping inlineImportAction =
        new StructuredMutationAction.ImportCustomXmlMapping(
            new CustomXmlImportInput(
                new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                TextSourceInput.inline("<root/>")));

    for (ChartPlotInput plot : standardInputChart.plots()) {
      assertTrue(SourceBackedInputRequirements.requiresStandardInput(plot));
    }
    assertTrue(SourceBackedInputRequirements.requiresStandardInput(standardInputChart));
    assertTrue(SourceBackedInputRequirements.requiresStandardInput(signatureLineWithStdIn()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(signatureLineWithInlineBinary()));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new SignatureLineInput(
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
            new CustomXmlImportInput(
                new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                TextSourceInput.standardInput())));
    assertFalse(
        SourceBackedInputRequirements.requiresStandardInput(
            new CustomXmlImportInput(
                new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                TextSourceInput.inline("<root/>"))));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new DrawingMutationAction.SetSignatureLine(signatureLineWithStdIn())));
    assertTrue(
        SourceBackedInputRequirements.requiresStandardInput(
            new StructuredMutationAction.ImportCustomXmlMapping(
                new CustomXmlImportInput(
                    new CustomXmlMappingLocator(1L, "CORSO_mapping"),
                    TextSourceInput.standardInput()))));

    WorkbookPlan resolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of(
                    new MutationStep("set-chart", new SheetSelector.ByName("Ops"), chartAction),
                    new MutationStep(
                        "set-signature", new SheetSelector.ByName("Ops"), signatureAction),
                    new MutationStep(
                        "import-custom-xml", new WorkbookSelector.Current(), inlineImportAction))),
            new ExecutionInputBindings(workingDirectory, "stdin".getBytes(StandardCharsets.UTF_8)));

    DrawingMutationAction.SetChart resolvedChartAction =
        assertInstanceOf(
            DrawingMutationAction.SetChart.class,
            assertInstanceOf(MutationStep.class, resolved.steps().get(0)).action());
    assertEquals(expectedInputPlotTypes(), plotTypes(resolvedChartAction.chart().plots()));
    for (ChartPlotInput plot : resolvedChartAction.chart().plots()) {
      ChartSeriesInput series = firstSeries(plot);
      ChartTitleInput.Text title = assertInstanceOf(ChartTitleInput.Text.class, series.title());
      assertEquals(
          "Resolved Series", assertInstanceOf(TextSourceInput.Inline.class, title.source()).text());
    }
    assertSame(
        signatureAction, assertInstanceOf(MutationStep.class, resolved.steps().get(1)).action());
    assertSame(
        inlineImportAction, assertInstanceOf(MutationStep.class, resolved.steps().get(2)).action());

    WorkbookPlan stableResolved =
        SourceBackedPlanResolver.resolve(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of(
                    new MutationStep(
                        "stable-chart", new SheetSelector.ByName("Ops"), stableChartAction),
                    new MutationStep(
                        "stable-signature",
                        new SheetSelector.ByName("Ops"),
                        stableSignatureAction))),
            new ExecutionInputBindings(workingDirectory, "stdin".getBytes(StandardCharsets.UTF_8)));

    assertSame(
        stableChartAction,
        assertInstanceOf(MutationStep.class, stableResolved.steps().get(0)).action());
    assertSame(
        stableSignatureAction,
        assertInstanceOf(MutationStep.class, stableResolved.steps().get(1)).action());
  }

  private static SignatureLineInput signatureLineWithStdIn() {
    return new SignatureLineInput(
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
            new PictureDataInput(ExcelPictureFormat.PNG, BinarySourceInput.standardInput())));
  }

  private static SignatureLineInput signatureLineWithInlineBinary() {
    return new SignatureLineInput(
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
            new PictureDataInput(
                ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64("cGF5bG9hZA=="))));
  }

  private static DrawingAnchorInput.TwoCell inputAnchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(1, 2, 0, 0),
        new DrawingMarkerInput(8, 14, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static ExcelDrawingAnchor.TwoCell excelAnchor() {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(1, 2, 0, 0),
        new ExcelDrawingMarker(8, 14, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static ChartSeriesInput inlineSeries(TextSourceInput title) {
    return new ChartSeriesInput(
        new ChartTitleInput.Text(title),
        new ChartDataSourceInput.StringLiteral(List.of("Jan", "Feb")),
        new ChartDataSourceInput.NumericLiteral(List.of(10.0d, 18.0d)),
        true,
        ExcelChartMarkerStyle.DIAMOND,
        (short) 6,
        4L);
  }

  private static ExcelChartSnapshot.Series snapshotSeries(ExcelChartSnapshot.Title title) {
    return new ExcelChartSnapshot.Series(
        title,
        new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan", "Feb")),
        new ExcelChartSnapshot.DataSource.NumericLiteral("0.0", List.of("10", "18")),
        true,
        ExcelChartMarkerStyle.DIAMOND,
        (short) 6,
        4L);
  }

  private static List<ChartPlotInput> inputPlots(ChartSeriesInput series) {
    return List.of(
        new ChartPlotInput.Area(false, ExcelChartGrouping.STANDARD, inputAxes(), List.of(series)),
        new ChartPlotInput.Area3D(
            false, ExcelChartGrouping.PERCENT_STACKED, 24, inputAxes(), List.of(series)),
        new ChartPlotInput.Bar(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            60,
            0,
            inputAxes(),
            List.of(series)),
        new ChartPlotInput.Bar3D(
            true,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.STACKED,
            32,
            88,
            ExcelChartBarShape.CONE,
            inputAxes(),
            List.of(series)),
        new ChartPlotInput.Doughnut(true, 30, 55, List.of(series)),
        new ChartPlotInput.Line(false, ExcelChartGrouping.STANDARD, inputAxes(), List.of(series)),
        new ChartPlotInput.Line3D(
            false, ExcelChartGrouping.STANDARD, 18, inputAxes(), List.of(series)),
        new ChartPlotInput.Pie(true, 90, List.of(series)),
        new ChartPlotInput.Pie3D(true, List.of(series)),
        new ChartPlotInput.Radar(false, ExcelChartRadarStyle.FILLED, inputAxes(), List.of(series)),
        new ChartPlotInput.Scatter(
            false, ExcelChartScatterStyle.SMOOTH_MARKER, scatterInputAxes(), List.of(series)),
        new ChartPlotInput.Surface(false, true, surfaceInputAxes(), List.of(series)),
        new ChartPlotInput.Surface3D(true, false, surfaceInputAxes(), List.of(series)));
  }

  private static List<ExcelChartSnapshot.Plot> snapshotPlots(ExcelChartSnapshot.Series series) {
    return List.of(
        new ExcelChartSnapshot.Area(
            false, ExcelChartGrouping.STANDARD, snapshotAxes(), List.of(series)),
        new ExcelChartSnapshot.Area3D(
            false, ExcelChartGrouping.PERCENT_STACKED, 24, snapshotAxes(), List.of(series)),
        new ExcelChartSnapshot.Bar(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            60,
            0,
            snapshotAxes(),
            List.of(series)),
        new ExcelChartSnapshot.Bar3D(
            true,
            ExcelChartBarDirection.BAR,
            ExcelChartBarGrouping.STACKED,
            32,
            88,
            ExcelChartBarShape.CONE,
            snapshotAxes(),
            List.of(series)),
        new ExcelChartSnapshot.Doughnut(true, 30, 55, List.of(series)),
        new ExcelChartSnapshot.Line(
            false, ExcelChartGrouping.STANDARD, snapshotAxes(), List.of(series)),
        new ExcelChartSnapshot.Line3D(
            false, ExcelChartGrouping.STANDARD, 18, snapshotAxes(), List.of(series)),
        new ExcelChartSnapshot.Pie(true, 90, List.of(series)),
        new ExcelChartSnapshot.Pie3D(true, List.of(series)),
        new ExcelChartSnapshot.Radar(
            false, ExcelChartRadarStyle.FILLED, snapshotAxes(), List.of(series)),
        new ExcelChartSnapshot.Scatter(
            false, ExcelChartScatterStyle.SMOOTH_MARKER, scatterSnapshotAxes(), List.of(series)),
        new ExcelChartSnapshot.Surface(false, true, surfaceSnapshotAxes(), List.of(series)),
        new ExcelChartSnapshot.Surface3D(true, false, surfaceSnapshotAxes(), List.of(series)));
  }

  private static List<ChartAxisInput> inputAxes() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisInput> scatterInputAxes() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartAxisInput> surfaceInputAxes() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartSnapshot.Axis> snapshotAxes() {
    return List.of(
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartSnapshot.Axis> scatterSnapshotAxes() {
    return List.of(
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartSnapshot.Axis> surfaceSnapshotAxes() {
    return List.of(
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartSnapshot.Axis(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<Class<?>> expectedDefinitionPlotTypes() {
    return List.of(
        ExcelChartDefinition.Area.class,
        ExcelChartDefinition.Area3D.class,
        ExcelChartDefinition.Bar.class,
        ExcelChartDefinition.Bar3D.class,
        ExcelChartDefinition.Doughnut.class,
        ExcelChartDefinition.Line.class,
        ExcelChartDefinition.Line3D.class,
        ExcelChartDefinition.Pie.class,
        ExcelChartDefinition.Pie3D.class,
        ExcelChartDefinition.Radar.class,
        ExcelChartDefinition.Scatter.class,
        ExcelChartDefinition.Surface.class,
        ExcelChartDefinition.Surface3D.class);
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
        ChartPlotInput.Area.class,
        ChartPlotInput.Area3D.class,
        ChartPlotInput.Bar.class,
        ChartPlotInput.Bar3D.class,
        ChartPlotInput.Doughnut.class,
        ChartPlotInput.Line.class,
        ChartPlotInput.Line3D.class,
        ChartPlotInput.Pie.class,
        ChartPlotInput.Pie3D.class,
        ChartPlotInput.Radar.class,
        ChartPlotInput.Scatter.class,
        ChartPlotInput.Surface.class,
        ChartPlotInput.Surface3D.class);
  }

  private static List<Class<?>> plotTypes(List<?> plots) {
    return plots.stream().<Class<?>>map(Object::getClass).toList();
  }

  private static ChartSeriesInput firstSeries(ChartPlotInput plot) {
    return switch (plot) {
      case ChartPlotInput.Area area -> area.series().getFirst();
      case ChartPlotInput.Area3D area3D -> area3D.series().getFirst();
      case ChartPlotInput.Bar bar -> bar.series().getFirst();
      case ChartPlotInput.Bar3D bar3D -> bar3D.series().getFirst();
      case ChartPlotInput.Doughnut doughnut -> doughnut.series().getFirst();
      case ChartPlotInput.Line line -> line.series().getFirst();
      case ChartPlotInput.Line3D line3D -> line3D.series().getFirst();
      case ChartPlotInput.Pie pie -> pie.series().getFirst();
      case ChartPlotInput.Pie3D pie3D -> pie3D.series().getFirst();
      case ChartPlotInput.Radar radar -> radar.series().getFirst();
      case ChartPlotInput.Scatter scatter -> scatter.series().getFirst();
      case ChartPlotInput.Surface surface -> surface.series().getFirst();
      case ChartPlotInput.Surface3D surface3D -> surface3D.series().getFirst();
    };
  }
}
