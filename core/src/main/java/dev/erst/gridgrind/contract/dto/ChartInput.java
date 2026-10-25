package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Authored chart input. One chart may contain one or more plots. */
public record ChartInput(
    String name,
    DrawingAnchorInput anchor,
    ChartTitleInput title,
    ChartLegendInput legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    List<ChartPlotInput> plots) {
  /** Reads one chart definition while applying the documented legend and display defaults. */
  @JsonCreator
  public ChartInput(
      @JsonProperty("name") String name,
      @JsonProperty("anchor") DrawingAnchorInput anchor,
      @JsonProperty("title") ChartTitleInput title,
      @JsonProperty("legend") ChartLegendInput legend,
      @JsonProperty("displayBlanksAs") ExcelChartDisplayBlanksAs displayBlanksAs,
      @JsonProperty("plotOnlyVisibleCells") Boolean plotOnlyVisibleCells,
      @JsonProperty("plots") List<ChartPlotInput> plots) {
    this(
        name,
        anchor,
        Objects.requireNonNull(title, "title must not be null"),
        Objects.requireNonNull(legend, "legend must not be null"),
        Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null"),
        Objects.requireNonNull(plotOnlyVisibleCells, "plotOnlyVisibleCells must not be null")
            .booleanValue(),
        plots);
  }

  /** Creates one chart with the standard visible legend and gap display state. */
  public static ChartInput withStandardDisplay(
      String name, DrawingAnchorInput anchor, List<ChartPlotInput> plots) {
    return new ChartInput(
        name,
        anchor,
        new ChartTitleInput.None(),
        new ChartLegendInput.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        plots);
  }

  public ChartInput {
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
