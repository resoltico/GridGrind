package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceInput;
import dev.erst.gridgrind.core.model.chart.ChartInput;
import dev.erst.gridgrind.core.model.chart.ChartLegendInput;
import dev.erst.gridgrind.core.model.chart.ChartPlotInput;
import dev.erst.gridgrind.core.model.chart.ChartSeriesInput;
import dev.erst.gridgrind.core.model.chart.ChartTitleInput;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorInput;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import java.util.List;
import java.util.Optional;

/** Shared chart input builders for shipped example workbook plans. */
final class ExampleChartInputs {
  private ExampleChartInputs() {}

  static ChartInput clusteredColumnComparisonChart(
      String chartName,
      DrawingAnchorInput.TwoCell anchor,
      String title,
      String categorySource,
      String firstSeriesTitle,
      String firstSeriesValues,
      String secondSeriesTitle,
      String secondSeriesValues) {
    return new ChartInput(
        chartName,
        anchor,
        new ChartTitleInput.Text(TextSourceInput.inline(title)),
        new ChartLegendInput.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(
            new ChartPlotInput.Bar(
                true,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.of(150),
                Optional.of(0),
                List.of(
                    series(firstSeriesTitle, categorySource, firstSeriesValues),
                    series(secondSeriesTitle, categorySource, secondSeriesValues)))));
  }

  private static ChartSeriesInput series(
      String title, String categorySource, String valueSourceReference) {
    return new ChartSeriesInput(
        new ChartTitleInput.Text(TextSourceInput.inline(title)),
        new ChartDataSourceInput.Reference(categorySource),
        new ChartDataSourceInput.Reference(valueSourceReference),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }
}
