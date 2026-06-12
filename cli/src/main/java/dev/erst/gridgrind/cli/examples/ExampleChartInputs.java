package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
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
