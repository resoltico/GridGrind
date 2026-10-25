package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import java.util.Objects;

/** Requested chart-legend state. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartLegendInput.Hidden.class, name = "HIDDEN"),
  @JsonSubTypes.Type(value = ChartLegendInput.Visible.class, name = "VISIBLE")
})
public sealed interface ChartLegendInput permits ChartLegendInput.Hidden, ChartLegendInput.Visible {
  /** Hide the legend entirely. */
  record Hidden() implements ChartLegendInput {}

  /** Show the legend at one position. */
  record Visible(ExcelChartLegendPosition position) implements ChartLegendInput {
    public Visible {
      Objects.requireNonNull(position, "position must not be null");
    }
  }
}
