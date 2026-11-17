package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.fixture.ChartInputFixtures;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarShape;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Extended chart plot preparation and absent signature preview operands. */
class ChartPreparationTest {
  @Test
  void preparesEveryExtendedChartPlotAndAbsentSignaturePreview() {
    ChartSpec chart =
        new ChartSpec(
            "ExtendedChart",
            ChartInputFixtures.anchor(),
            new ChartTitleSpec.Text(TextSourceSpec.inline("Extended")),
            new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP_RIGHT),
            ExcelChartDisplayBlanksAs.SPAN,
            false,
            ChartInputFixtures.plots(ChartInputFixtures.series(TextSourceSpec.inline("Series"))));
    ResolvedChart definition = MutationPreparation.toResolvedChart(chart);
    ResolvedSignatureLine signatureLine =
        MutationPreparation.toResolvedSignatureLine(
            new SignatureLineSpec(
                "OpsSignature",
                ChartInputFixtures.anchor(),
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
    assertEquals(Optional.empty(), signatureLine.plainSignatureFormat());
    assertEquals(Optional.empty(), signatureLine.plainSignature());
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

  private static List<Class<?>> plotTypes(List<?> plots) {
    return plots.stream().<Class<?>>map(Object::getClass).toList();
  }
}
