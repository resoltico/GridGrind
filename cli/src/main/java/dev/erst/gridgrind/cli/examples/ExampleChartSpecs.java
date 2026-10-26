package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import java.util.List;
import java.util.Optional;

/** Shared chart input builders for shipped example workbook plans. */
final class ExampleChartSpecs {
  private ExampleChartSpecs() {}

  static ChartSpec clusteredColumnComparisonChart(
      String chartName,
      DrawingAnchorSpec.TwoCell anchor,
      String title,
      String categorySource,
      String firstSeriesTitle,
      String firstSeriesValues,
      String secondSeriesTitle,
      String secondSeriesValues) {
    return new ChartSpec(
        chartName,
        anchor,
        new ChartTitleSpec.Text(TextSourceSpec.inline(title)),
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(
            new ChartPlotSpec.Bar(
                true,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.of(150),
                Optional.of(0),
                List.of(
                    series(firstSeriesTitle, categorySource, firstSeriesValues),
                    series(secondSeriesTitle, categorySource, secondSeriesValues)))));
  }

  private static ChartSeriesSpec series(
      String title, String categorySource, String valueSourceReference) {
    return new ChartSeriesSpec(
        new ChartTitleSpec.Text(TextSourceSpec.inline(title)),
        new ChartDataSourceSpec.Reference(categorySource),
        new ChartDataSourceSpec.Reference(valueSourceReference),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }
}
