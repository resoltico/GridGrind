package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import java.util.Objects;

/** Requested chart-legend state. */
@Union(discriminator = "type", catalog = true, displayOrder = 51)
public sealed interface ChartLegendInput
    permits dev.erst.gridgrind.core.model.chart.ChartLegendInput.Hidden,
        dev.erst.gridgrind.core.model.chart.ChartLegendInput.Visible {
  /** Hide the legend entirely. */
  @Variant(id = "HIDDEN", summary = "Hide the legend entirely.")
  record Hidden() implements ChartLegendInput {}

  /** Show the legend at one position. */
  @Variant(id = "VISIBLE", summary = "Show the legend at one explicit position.")
  record Visible(ExcelChartLegendPosition position) implements ChartLegendInput {
    public Visible {
      Objects.requireNonNull(position, "position must not be null");
    }
  }
}
