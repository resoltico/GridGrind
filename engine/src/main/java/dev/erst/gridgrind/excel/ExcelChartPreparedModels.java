package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

record ResolvedAreaReference(XSSFSheet sheet, AreaReference areaReference) {
  ResolvedAreaReference {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(areaReference, "areaReference must not be null");
  }
}

record ResolvedChartSource(
    String referenceFormula,
    XSSFSheet sheet,
    AreaReference areaReference,
    boolean numeric,
    List<String> stringValues,
    List<Double> numericValues) {
  ResolvedChartSource {
    Objects.requireNonNull(referenceFormula, "referenceFormula must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(areaReference, "areaReference must not be null");
    stringValues =
        List.copyOf(Objects.requireNonNull(stringValues, "stringValues must not be null"));
    numericValues =
        List.copyOf(Objects.requireNonNull(numericValues, "numericValues must not be null"));
  }
}

/** Prepared chart payload validated fully before any chart mutation starts. */
sealed interface PreparedChartDefinition
    permits PreparedBarChart, PreparedLineChart, PreparedPieChart {
  /** Returns the target chart name. */
  String name();

  /** Returns the target chart anchor. */
  ExcelDrawingAnchor.TwoCell anchor();

  /** Returns the validated chart title payload. */
  PreparedChartTitle title();

  /** Returns the validated legend payload. */
  ExcelChartDefinition.Legend legend();

  /** Returns the validated blank-cell display behavior. */
  ExcelChartDisplayBlanksAs displayBlanksAs();

  /** Returns whether hidden cells stay excluded from the chart. */
  boolean plotOnlyVisibleCells();
}

/** Prepared chart-title payload validated before chart creation or mutation. */
sealed interface PreparedChartTitle
    permits PreparedChartTitleNone, PreparedChartTitleText, PreparedChartTitleFormula {}

/** Prepared series-title payload validated before chart creation or mutation. */
sealed interface PreparedSeriesTitle
    permits PreparedSeriesTitleNone, PreparedSeriesTitleText, PreparedSeriesTitleFormula {}

record PreparedBarChart(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
    PreparedChartTitle title,
    ExcelChartDefinition.Legend legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    boolean varyColors,
    ExcelChartBarDirection barDirection,
    List<PreparedChartSeries> series)
    implements PreparedChartDefinition {
  PreparedBarChart {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    Objects.requireNonNull(barDirection, "barDirection must not be null");
    series = List.copyOf(Objects.requireNonNull(series, "series must not be null"));
  }
}

record PreparedLineChart(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
    PreparedChartTitle title,
    ExcelChartDefinition.Legend legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    boolean varyColors,
    List<PreparedChartSeries> series)
    implements PreparedChartDefinition {
  PreparedLineChart {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    series = List.copyOf(Objects.requireNonNull(series, "series must not be null"));
  }
}

record PreparedPieChart(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
    PreparedChartTitle title,
    ExcelChartDefinition.Legend legend,
    ExcelChartDisplayBlanksAs displayBlanksAs,
    boolean plotOnlyVisibleCells,
    boolean varyColors,
    Integer firstSliceAngle,
    List<PreparedChartSeries> series)
    implements PreparedChartDefinition {
  PreparedPieChart {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(legend, "legend must not be null");
    Objects.requireNonNull(displayBlanksAs, "displayBlanksAs must not be null");
    series = List.copyOf(Objects.requireNonNull(series, "series must not be null"));
  }
}

record PreparedChartTitleNone() implements PreparedChartTitle {}

record PreparedChartTitleText(String text) implements PreparedChartTitle {
  PreparedChartTitleText {
    Objects.requireNonNull(text, "text must not be null");
  }
}

record PreparedChartTitleFormula(String cachedText, CellReference reference)
    implements PreparedChartTitle {
  PreparedChartTitleFormula {
    Objects.requireNonNull(cachedText, "cachedText must not be null");
    Objects.requireNonNull(reference, "reference must not be null");
  }
}

record PreparedSeriesTitleNone() implements PreparedSeriesTitle {}

record PreparedSeriesTitleText(String text) implements PreparedSeriesTitle {
  PreparedSeriesTitleText {
    Objects.requireNonNull(text, "text must not be null");
  }
}

record PreparedSeriesTitleFormula(String cachedText, CellReference reference)
    implements PreparedSeriesTitle {
  PreparedSeriesTitleFormula {
    Objects.requireNonNull(cachedText, "cachedText must not be null");
    Objects.requireNonNull(reference, "reference must not be null");
  }
}

record PreparedChartSeries(
    PreparedSeriesTitle title,
    org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> categories,
    org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<? extends Number> values) {
  PreparedChartSeries {
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(categories, "categories must not be null");
    Objects.requireNonNull(values, "values must not be null");
  }
}
