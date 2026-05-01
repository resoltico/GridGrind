package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Owns chart-series snapshot extraction shared across chart plot families. */
final class ExcelChartSeriesSnapshotSupport {
  private ExcelChartSeriesSnapshotSupport() {}

  static List<ExcelChartSnapshot.Series> snapshotAreaSeries(
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData data,
      ExcelFormulaRuntime formulaRuntime) {
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
      XSSFSheet contextSheet,
      XDDFChartData.Series series,
      CTSerTx title,
      Boolean smooth,
      ExcelChartMarkerStyle markerStyle,
      Short markerSize,
      ExcelFormulaRuntime formulaRuntime) {
    return snapshotSeries(
        contextSheet, series, title, smooth, markerStyle, markerSize, null, formulaRuntime);
  }

  private static ExcelChartSnapshot.Series snapshotSeries(
      XSSFSheet contextSheet,
      XDDFChartData.Series series,
      CTSerTx title,
      Boolean smooth,
      ExcelChartMarkerStyle markerStyle,
      Short markerSize,
      Long explosion,
      ExcelFormulaRuntime formulaRuntime) {
    return new ExcelChartSnapshot.Series(
        ExcelChartSnapshotSupport.snapshotSeriesTitle(contextSheet, title, formulaRuntime),
        snapshotDataSource(contextSheet, series.getCategoryData(), formulaRuntime),
        snapshotDataSource(contextSheet, series.getValuesData(), formulaRuntime),
        smooth,
        markerStyle,
        markerSize,
        explosion);
  }

  private static Boolean smooth(boolean present, boolean value) {
    return present ? value : null;
  }

  static ExcelChartSnapshot.DataSource snapshotDataSource(
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      ExcelFormulaRuntime formulaRuntime) {
    if (source == null) {
      throw new IllegalStateException("Chart series is missing its data source");
    }
    if (source.isReference()) {
      String referenceFormula = source.getDataRangeReference();
      List<String> values =
          resolvedOrCachedReferenceValues(contextSheet, referenceFormula, source, formulaRuntime);
      return source.isNumeric()
          ? new ExcelChartSnapshot.DataSource.NumericReference(
              referenceFormula, source.getFormatCode(), values)
          : new ExcelChartSnapshot.DataSource.StringReference(referenceFormula, values);
    }
    List<String> values = cachedPointValues(source);
    return source.isNumeric()
        ? new ExcelChartSnapshot.DataSource.NumericLiteral(source.getFormatCode(), values)
        : new ExcelChartSnapshot.DataSource.StringLiteral(values);
  }

  static List<String> resolvedOrCachedReferenceValues(
      XSSFSheet contextSheet,
      String referenceFormula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      ExcelFormulaRuntime formulaRuntime) {
    if (referenceFormula != null && !referenceFormula.isBlank()) {
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

  static Short markerSize(CTMarker marker) {
    return marker != null && marker.isSetSize() ? (short) marker.getSize().getVal() : null;
  }
}
