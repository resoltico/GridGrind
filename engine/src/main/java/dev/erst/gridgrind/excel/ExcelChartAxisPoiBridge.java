package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import java.util.Objects;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDateAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFSeriesAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;

/** Maps chart axis families and creation requests between GridGrind and POI. */
final class ExcelChartAxisPoiBridge {
  private ExcelChartAxisPoiBridge() {}

  static ExcelChartAxisKind axisKind(XDDFChartAxis axis) {
    Objects.requireNonNull(axis, "axis must not be null");
    return switch (axis) {
      case XDDFCategoryAxis _ -> ExcelChartAxisKind.CATEGORY;
      case XDDFDateAxis _ -> ExcelChartAxisKind.DATE;
      case XDDFSeriesAxis _ -> ExcelChartAxisKind.SERIES;
      case XDDFValueAxis _ -> ExcelChartAxisKind.VALUE;
      default ->
          throw new IllegalArgumentException("Unsupported chart axis family: " + axis.getClass());
    };
  }

  static XDDFChartAxis createAxis(
      XDDFChart chart, ExcelChartAxisKind axisKind, ExcelChartAxisPosition position) {
    Objects.requireNonNull(chart, "chart must not be null");
    Objects.requireNonNull(axisKind, "axisKind must not be null");
    Objects.requireNonNull(position, "position must not be null");
    return switch (axisKind) {
      case CATEGORY -> chart.createCategoryAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
      case DATE -> chart.createDateAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
      case SERIES -> chart.createSeriesAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
      case VALUE -> chart.createValueAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
    };
  }
}
