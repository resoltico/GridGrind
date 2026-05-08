package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Owns chart-series snapshot extraction shared across chart plot families. */
final class ExcelChartSeriesSnapshotSupport {
  private ExcelChartSeriesSnapshotSupport() {}

  static List<ExcelChartSnapshot.Series> snapshotAreaSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTAreaSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotArea3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTAreaSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotBarSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTBarSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotBar3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet, value, value.getCTBarSer().getTx(), null, null, null, formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotDoughnutSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTPieSer().getTx(),
              null,
              null,
              null,
              value.getExplosion(),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotLineSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTLineSer().getTx(),
              smooth(value.getCTLineSer().isSetSmooth(), value.isSmooth()),
              markerStyle(value.getCTLineSer().getMarker()).orElse(null),
              markerSize(value.getCTLineSer().getMarker()),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotLine3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTLineSer().getTx(),
              smooth(value.getCTLineSer().isSetSmooth(), value.isSmooth()),
              markerStyle(value.getCTLineSer().getMarker()).orElse(null),
              markerSize(value.getCTLineSer().getMarker()),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotPieSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTPieSer().getTx(),
              null,
              null,
              null,
              value.getExplosion(),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotPie3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTPieSer().getTx(),
              null,
              null,
              null,
              value.getExplosion(),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotRadarSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTRadarSer().getTx(),
              null,
              null,
              null,
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotScatterSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTScatterSer().getTx(),
              smooth(value.getCTScatterSer().isSetSmooth(), value.isSmooth()),
              markerStyle(value.getCTScatterSer().getMarker()).orElse(null),
              markerSize(value.getCTScatterSer().getMarker()),
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotSurfaceSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTSurfaceSer().getTx(),
              null,
              null,
              null,
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartSnapshot.Series> snapshotSurface3DSeries(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData data,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      var value =
          (org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData.Series) data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              value.getCTSurfaceSer().getTx(),
              null,
              null,
              null,
              formulaRuntime));
    }
    return List.copyOf(series);
  }

  private static ExcelChartSnapshot.Series snapshotSeries(
      @Nullable XSSFSheet contextSheet,
      XDDFChartData.Series series,
      CTSerTx title,
      @Nullable Boolean smooth,
      @Nullable ExcelChartMarkerStyle markerStyle,
      @Nullable Short markerSize,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return snapshotSeries(
        contextSheet, series, title, smooth, markerStyle, markerSize, null, formulaRuntime);
  }

  private static ExcelChartSnapshot.Series snapshotSeries(
      @Nullable XSSFSheet contextSheet,
      XDDFChartData.Series series,
      CTSerTx title,
      @Nullable Boolean smooth,
      @Nullable ExcelChartMarkerStyle markerStyle,
      @Nullable Short markerSize,
      @Nullable Long explosion,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return new ExcelChartSnapshot.Series(
        ExcelChartSnapshotSupport.snapshotSeriesTitle(contextSheet, title, formulaRuntime),
        snapshotDataSource(contextSheet, series.getCategoryData(), formulaRuntime),
        snapshotDataSource(contextSheet, series.getValuesData(), formulaRuntime),
        Optional.ofNullable(smooth),
        Optional.ofNullable(markerStyle),
        Optional.ofNullable(markerSize),
        Optional.ofNullable(explosion));
  }

  private static @Nullable Boolean smooth(boolean present, boolean value) {
    return present ? value : null;
  }

  static ExcelChartSnapshot.DataSource snapshotDataSource(
      @Nullable XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    if (source == null) {
      throw new IllegalStateException("Chart series is missing its data source");
    }
    if (source.isReference()) {
      String referenceFormula = source.getDataRangeReference();
      List<String> values =
          resolvedOrCachedReferenceValues(contextSheet, referenceFormula, source, formulaRuntime);
      return source.isNumeric()
          ? new ExcelChartSnapshot.DataSource.NumericReference(
              referenceFormula, Optional.ofNullable(source.getFormatCode()), values)
          : new ExcelChartSnapshot.DataSource.StringReference(referenceFormula, values);
    }
    List<String> values = cachedPointValues(source);
    return source.isNumeric()
        ? new ExcelChartSnapshot.DataSource.NumericLiteral(
            Optional.ofNullable(source.getFormatCode()), values)
        : new ExcelChartSnapshot.DataSource.StringLiteral(values);
  }

  static List<String> resolvedOrCachedReferenceValues(
      @Nullable XSSFSheet contextSheet,
      String referenceFormula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    if (contextSheet != null && referenceFormula != null && !referenceFormula.isBlank()) {
      try {
        return ExcelChartSourceSupport.resolveChartSource(
                contextSheet, referenceFormula, formulaRuntime)
            .stringValues();
      } catch (RuntimeException ignored) {
        // Fall back to the embedded chart cache when the reference cannot be resolved.
      }
    }
    return cachedPointValues(source);
  }

  private static List<String> cachedPointValues(
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    List<String> values = new ArrayList<>();
    for (int index = 0; index < source.getPointCount(); index++) {
      Object point;
      try {
        point = source.getPointAt(index);
      } catch (IndexOutOfBoundsException exception) {
        point = null;
      }
      values.add(point == null ? "" : point.toString());
    }
    return List.copyOf(values);
  }

  static Optional<ExcelChartMarkerStyle> markerStyle(CTMarker marker) {
    if (marker == null || !marker.isSetSymbol()) {
      return Optional.empty();
    }
    return Optional.ofNullable(
        switch (marker.getSymbol().getVal().toString().toUpperCase(java.util.Locale.ROOT)) {
          case "CIRCLE" -> ExcelChartMarkerStyle.CIRCLE;
          case "DASH" -> ExcelChartMarkerStyle.DASH;
          case "DIAMOND" -> ExcelChartMarkerStyle.DIAMOND;
          case "DOT" -> ExcelChartMarkerStyle.DOT;
          case "NONE" -> ExcelChartMarkerStyle.NONE;
          case "PICTURE" -> ExcelChartMarkerStyle.PICTURE;
          case "PLUS" -> ExcelChartMarkerStyle.PLUS;
          case "SQUARE" -> ExcelChartMarkerStyle.SQUARE;
          case "STAR" -> ExcelChartMarkerStyle.STAR;
          case "TRIANGLE" -> ExcelChartMarkerStyle.TRIANGLE;
          case "X" -> ExcelChartMarkerStyle.X;
          default -> null;
        });
  }

  static @Nullable Short markerSize(CTMarker marker) {
    return marker != null && marker.isSetSize() ? (short) marker.getSize().getVal() : null;
  }
}
