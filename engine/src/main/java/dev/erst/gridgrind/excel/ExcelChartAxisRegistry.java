package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFSeriesAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;

/** Thread-confined registry that creates and reuses POI axes for one authored chart. */
final class ExcelChartAxisRegistry {
  private final XDDFChart chart;
  private final Map<ExcelChartDefinition.Axis, XDDFChartAxis> axesByDefinition =
      new ConcurrentHashMap<>();

  ExcelChartAxisRegistry(XDDFChart chart) {
    this.chart = Objects.requireNonNull(chart, "chart must not be null");
  }

  CategoryValueAxes categoryValueAxes(List<ExcelChartDefinition.Axis> definitions) {
    XDDFChartAxis categoryAxis = null;
    XDDFValueAxis valueAxis = null;
    for (ExcelChartDefinition.Axis definition : definitions) {
      XDDFChartAxis axis = axis(definition);
      if (definition.kind().isCategoryAxis()) {
        categoryAxis = axis;
      }
      if (definition.kind().isValueAxis()) {
        valueAxis = (XDDFValueAxis) axis;
      }
      if (definition.kind().isSeriesAxis()) {
        throw new IllegalArgumentException("Series axis is unsupported for this plot family");
      }
    }
    if (categoryAxis == null || valueAxis == null) {
      throw new IllegalArgumentException(
          "Plot must declare one category/date axis and one value axis");
    }
    crossAxes(List.of(categoryAxis, valueAxis));
    return new CategoryValueAxes(categoryAxis, valueAxis);
  }

  ScatterAxes scatterAxes(List<ExcelChartDefinition.Axis> definitions) {
    List<XDDFValueAxis> valueAxes =
        definitions.stream().map(this::axis).map(axis -> (XDDFValueAxis) axis).toList();
    if (valueAxes.size() != 2) {
      throw new IllegalArgumentException("Scatter plots must declare exactly two value axes");
    }
    crossAxes(List.copyOf(valueAxes));
    return new ScatterAxes(valueAxes.getFirst(), valueAxes.get(1));
  }

  SurfaceAxes surfaceAxes(List<ExcelChartDefinition.Axis> definitions) {
    XDDFChartAxis categoryAxis = null;
    XDDFValueAxis valueAxis = null;
    XDDFSeriesAxis seriesAxis = null;
    for (ExcelChartDefinition.Axis definition : definitions) {
      XDDFChartAxis axis = axis(definition);
      if (definition.kind().isCategoryAxis()) {
        categoryAxis = axis;
      }
      if (definition.kind().isValueAxis()) {
        valueAxis = (XDDFValueAxis) axis;
      }
      if (definition.kind().isSeriesAxis()) {
        seriesAxis = (XDDFSeriesAxis) axis;
      }
    }
    if (categoryAxis == null || valueAxis == null || seriesAxis == null) {
      throw new IllegalArgumentException(
          "Surface plots must declare one category/date axis, one value axis, and one series axis");
    }
    crossAxes(List.of(categoryAxis, valueAxis, seriesAxis));
    return new SurfaceAxes(categoryAxis, valueAxis, seriesAxis);
  }

  private XDDFChartAxis axis(ExcelChartDefinition.Axis definition) {
    XDDFChartAxis existing = axesByDefinition.get(definition);
    if (existing != null) {
      return existing;
    }
    XDDFChartAxis created =
        ExcelChartPoiBridge.createAxis(chart, definition.kind(), definition.position());
    created.setCrosses(ExcelChartPoiBridge.toPoiAxisCrosses(definition.crosses()));
    created.setVisible(definition.visible());
    axesByDefinition.put(definition, created);
    return created;
  }

  private static void crossAxes(List<? extends XDDFChartAxis> axes) {
    for (int index = 0; index < axes.size(); index++) {
      for (int other = index + 1; other < axes.size(); other++) {
        axes.get(index).crossAxis(axes.get(other));
        axes.get(other).crossAxis(axes.get(index));
      }
    }
  }

  /** Pair of category/date axis and value axis for standard chart families. */
  record CategoryValueAxes(XDDFChartAxis categoryAxis, XDDFValueAxis valueAxis) {}

  /** Pair of value axes for scatter charts. */
  record ScatterAxes(XDDFValueAxis xAxis, XDDFValueAxis yAxis) {}

  /** Triple of category/date, value, and series axes for surface charts. */
  record SurfaceAxes(
      XDDFChartAxis categoryAxis, XDDFValueAxis valueAxis, XDDFSeriesAxis seriesAxis) {}
}
