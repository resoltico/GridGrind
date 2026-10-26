package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Authored chart input. One chart may contain one or more plots. */
@CatalogType(
    id = "ChartSpec",
    summary =
        "One authored chart with an explicit drawing anchor, chart-level presentation state, and one or more plots. Use ChartSpec.withStandardDisplay(...) from Java authoring when the standard legend and blank-cell display policy is intended.",
    displayOrder = 33)
public record ChartSpec(
    String name,
    DrawingAnchorSpec anchor,
    ChartTitleSpec title,
    ChartLegendSpec legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    List<ChartPlotSpec> plots) {

  /** Creates one chart with the standard visible legend and gap display state. */
  public static ChartSpec withStandardDisplay(
      String name, DrawingAnchorSpec anchor, List<ChartPlotSpec> plots) {
    return new ChartSpec(
        name,
        anchor,
        new ChartTitleSpec.None(),
        new ChartLegendSpec.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        plots);
  }

  public ChartSpec {
    name = requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    plots = copyNonEmptyValues(plots, "plots");
  }

  static <T> List<T> copyNonEmptyValues(List<T> values, String fieldName) {
    List<T> copiedValues = copyValues(values, fieldName);
    if (copiedValues.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return copiedValues;
  }

  static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copiedValues = new ArrayList<>(values.size());
    for (T value : values) {
      copiedValues.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copiedValues);
  }

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
