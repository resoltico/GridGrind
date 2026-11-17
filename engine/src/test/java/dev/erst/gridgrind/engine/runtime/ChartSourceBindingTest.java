package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.fixture.ChartInputFixtures;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.*;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Checks file/stdin chart operands and unchanged source-free actions through source binding. */
class ChartSourceBindingTest {

  @Test
  void sourceBackedHelpersCoverExtendedChartPlotFamiliesAndStableNullBranches() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-chart-surface-");
    Files.writeString(
        workingDirectory.resolve("series-title.txt"), "Resolved Series", StandardCharsets.UTF_8);

    ChartSpec fileBackedChart =
        new ChartSpec(
            "ExtendedChart",
            ChartInputFixtures.anchor(),
            new ChartTitleSpec.Text(TextSourceSpec.inline("Extended")),
            new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            ChartInputFixtures.plots(
                ChartInputFixtures.series(TextSourceSpec.utf8File("series-title.txt"))));
    ChartSpec standardInputChart =
        new ChartSpec(
            "ExtendedChart",
            ChartInputFixtures.anchor(),
            new ChartTitleSpec.None(),
            new ChartLegendSpec.Hidden(),
            ExcelChartDisplayBlanksAs.ZERO,
            false,
            ChartInputFixtures.plots(ChartInputFixtures.series(TextSourceSpec.standardInput())));
    DrawingMutationAction.SetChart chartAction =
        new DrawingMutationAction.SetChart(fileBackedChart);
    DrawingMutationAction.SetSignatureLine signatureAction =
        new DrawingMutationAction.SetSignatureLine(
            new SignatureLineSpec(
                "OpsSignature",
                ChartInputFixtures.anchor(),
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
                ChartInputFixtures.anchor(),
                new ChartTitleSpec.None(),
                new ChartLegendSpec.Hidden(),
                ExcelChartDisplayBlanksAs.GAP,
                true,
                ChartInputFixtures.plots(
                    ChartInputFixtures.series(TextSourceSpec.inline("Stable Series")))));
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
                ChartInputFixtures.anchor(),
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
        ChartInputFixtures.anchor(),
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
        ChartInputFixtures.anchor(),
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
