package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import java.util.Objects;

/** Requested chart-legend state. */
@Union(discriminator = "type", catalog = true, displayOrder = 51)
public sealed interface ChartLegendSpec
    permits dev.erst.gridgrind.core.model.chart.ChartLegendSpec.Hidden,
        dev.erst.gridgrind.core.model.chart.ChartLegendSpec.Visible {
  /** Hide the legend entirely. */
  @Variant(id = "HIDDEN", summary = "Hide the legend entirely.")
  record Hidden() implements ChartLegendSpec {}

  /** Show the legend at one position. */
  @Variant(id = "VISIBLE", summary = "Show the legend at one explicit position.")
  record Visible(ExcelChartLegendPosition position) implements ChartLegendSpec {
    public Visible {
      Objects.requireNonNull(position, "position must not be null");
    }
  }
}
