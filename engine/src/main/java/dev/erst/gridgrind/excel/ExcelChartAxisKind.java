package dev.erst.gridgrind.excel;

import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;

/** Supported chart-axis families. */
public enum ExcelChartAxisKind {
  CATEGORY {
    @Override
    XDDFChartAxis createAxis(XDDFChart chart, ExcelChartAxisPosition position) {
      return chart.createCategoryAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
    }

    @Override
    boolean isCategoryAxis() {
      return true;
    }
  },
  DATE {
    @Override
    XDDFChartAxis createAxis(XDDFChart chart, ExcelChartAxisPosition position) {
      return chart.createDateAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
    }

    @Override
    boolean isCategoryAxis() {
      return true;
    }
  },
  SERIES {
    @Override
    XDDFChartAxis createAxis(XDDFChart chart, ExcelChartAxisPosition position) {
      return chart.createSeriesAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
    }

    @Override
    boolean isSeriesAxis() {
      return true;
    }
  },
  VALUE {
    @Override
    XDDFChartAxis createAxis(XDDFChart chart, ExcelChartAxisPosition position) {
      return chart.createValueAxis(ExcelChartPoiBridge.toPoiAxisPosition(position));
    }

    @Override
    boolean isValueAxis() {
      return true;
    }
  };

  abstract XDDFChartAxis createAxis(XDDFChart chart, ExcelChartAxisPosition position);

  boolean isCategoryAxis() {
    return false;
  }

  boolean isSeriesAxis() {
    return false;
  }

  boolean isValueAxis() {
    return false;
  }
}
