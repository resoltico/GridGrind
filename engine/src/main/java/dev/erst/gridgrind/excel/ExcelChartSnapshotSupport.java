package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Chart snapshot and chart-readback helpers. */
final class ExcelChartSnapshotSupport {
  private ExcelChartSnapshotSupport() {}

  static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    ExcelChartSnapshot snapshot = snapshotChart(chart, graphicFrame);
    return new ExcelDrawingObjectSnapshot.Chart(
        resolvedChartName(graphicFrame),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
        !(snapshot instanceof ExcelChartSnapshot.Unsupported),
        chartPlotTypeTokens(chart),
        titleSummary(snapshot));
  }

  static ExcelChartSnapshot snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    List<org.apache.poi.xddf.usermodel.chart.XDDFChartData> chartData = chart.getChartSeries();
    List<String> plotTypeTokens = chartPlotTypeTokens(chartData);
    if (chartData.size() != 1) {
      return new ExcelChartSnapshot.Unsupported(
          resolvedChartName(graphicFrame),
          ExcelDrawingAnchorSupport.snapshotAnchor(
              ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
          plotTypeTokens,
          "Only single-plot simple charts are modeled authoritatively.");
    }

    org.apache.poi.xddf.usermodel.chart.XDDFChartData data = chartData.getFirst();
    try {
      return switch (data) {
        case org.apache.poi.xddf.usermodel.chart.XDDFBarChartData barChartData ->
            new ExcelChartSnapshot.Bar(
                resolvedChartName(graphicFrame),
                ExcelDrawingAnchorSupport.snapshotAnchor(
                    ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
                snapshotTitle(chart),
                snapshotLegend(chart),
                snapshotDisplayBlanks(chart),
                chart.isPlotOnlyVisibleCells(),
                barVaryColors(chart),
                ExcelChartPoiBridge.fromPoiBarDirection(barChartData.getBarDirection()),
                snapshotAxes(chart),
                snapshotSeries(chart.getGraphicFrame().getDrawing().getSheet(), barChartData));
        case org.apache.poi.xddf.usermodel.chart.XDDFLineChartData lineChartData ->
            new ExcelChartSnapshot.Line(
                resolvedChartName(graphicFrame),
                ExcelDrawingAnchorSupport.snapshotAnchor(
                    ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
                snapshotTitle(chart),
                snapshotLegend(chart),
                snapshotDisplayBlanks(chart),
                chart.isPlotOnlyVisibleCells(),
                lineVaryColors(chart),
                snapshotAxes(chart),
                snapshotSeries(chart.getGraphicFrame().getDrawing().getSheet(), lineChartData));
        case org.apache.poi.xddf.usermodel.chart.XDDFPieChartData pieChartData ->
            new ExcelChartSnapshot.Pie(
                resolvedChartName(graphicFrame),
                ExcelDrawingAnchorSupport.snapshotAnchor(
                    ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
                snapshotTitle(chart),
                snapshotLegend(chart),
                snapshotDisplayBlanks(chart),
                chart.isPlotOnlyVisibleCells(),
                pieVaryColors(chart),
                pieChartData.getFirstSliceAngle(),
                snapshotSeries(chart.getGraphicFrame().getDrawing().getSheet(), pieChartData));
        default ->
            new ExcelChartSnapshot.Unsupported(
                resolvedChartName(graphicFrame),
                ExcelDrawingAnchorSupport.snapshotAnchor(
                    ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
                plotTypeTokens,
                "Chart plot family is outside the current modeled simple-chart contract.");
      };
    } catch (IllegalArgumentException | IllegalStateException exception) {
      return new ExcelChartSnapshot.Unsupported(
          resolvedChartName(graphicFrame),
          ExcelDrawingAnchorSupport.snapshotAnchor(
              ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
          plotTypeTokens,
          exception.getMessage());
    }
  }

  static XSSFChart chartForGraphicFrame(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    for (XSSFChart chart : drawing.getCharts()) {
      if (chart.getGraphicFrame().getId() == graphicFrame.getId()) {
        return chart;
      }
    }
    return null;
  }

  static ExcelChartSnapshot.Title snapshotTitle(XSSFChart chart) {
    if (!chart.getCTChart().isSetTitle()) {
      return new ExcelChartSnapshot.Title.None();
    }
    String formula = chart.getTitleFormula();
    if (formula != null) {
      return new ExcelChartSnapshot.Title.Formula(formula, cachedTitleText(chart, formula));
    }
    String text = chart.getTitleText().getString();
    return text.isBlank()
        ? new ExcelChartSnapshot.Title.None()
        : new ExcelChartSnapshot.Title.Text(text);
  }

  static String cachedTitleText(XSSFChart chart, String formula) {
    if (!chart.getCTChart().isSetTitle()
        || !chart.getCTChart().getTitle().isSetTx()
        || !chart.getCTChart().getTitle().getTx().isSetStrRef()
        || !chart.getCTChart().getTitle().getTx().getStrRef().isSetStrCache()
        || chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().sizeOfPtArray() == 0) {
      return resolvedTitleFormulaText(chart, formula);
    }
    String cachedText =
        chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV();
    return cachedText.isBlank() ? resolvedTitleFormulaText(chart, formula) : cachedText;
  }

  static String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    try {
      XSSFSheet sheet = chart.getGraphicFrame().getDrawing().getSheet();
      return ExcelChartSourceSupport.scalarText(
          sheet,
          ExcelChartSourceSupport.resolveSingleCellReference(
              sheet, formula, "Chart title formula"));
    } catch (RuntimeException exception) {
      return "";
    }
  }

  static boolean barVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfBarChartArray() > 0
        && chart.getCTChart().getPlotArea().getBarChartArray(0).isSetVaryColors()
        && chart.getCTChart().getPlotArea().getBarChartArray(0).getVaryColors().getVal();
  }

  static boolean lineVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfLineChartArray() > 0
        && chart.getCTChart().getPlotArea().getLineChartArray(0).isSetVaryColors()
        && chart.getCTChart().getPlotArea().getLineChartArray(0).getVaryColors().getVal();
  }

  static boolean pieVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfPieChartArray() > 0
        && chart.getCTChart().getPlotArea().getPieChartArray(0).isSetVaryColors()
        && chart.getCTChart().getPlotArea().getPieChartArray(0).getVaryColors().getVal();
  }

  static ExcelChartSnapshot.Title snapshotSeriesTitle(
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title) {
    if (title == null) {
      return new ExcelChartSnapshot.Title.None();
    }
    if (title.isSetStrRef()) {
      String cachedText =
          title.getStrRef().isSetStrCache() && title.getStrRef().getStrCache().sizeOfPtArray() > 0
              ? title.getStrRef().getStrCache().getPtArray(0).getV()
              : "";
      return new ExcelChartSnapshot.Title.Formula(title.getStrRef().getF(), cachedText);
    }
    return title.isSetV()
        ? new ExcelChartSnapshot.Title.Text(title.getV())
        : new ExcelChartSnapshot.Title.None();
  }

  static String titleSummary(ExcelChartSnapshot.Title title) {
    return switch (title) {
      case ExcelChartSnapshot.Title.None _ -> "";
      case ExcelChartSnapshot.Title.Text text -> text.text();
      case ExcelChartSnapshot.Title.Formula formula ->
          formula.cachedText().isEmpty() ? formula.formula() : formula.cachedText();
    };
  }

  private static String resolvedChartName(
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    String name = ExcelChartSourceSupport.nullIfBlank(graphicFrame.getName());
    return name != null ? name : "Chart-" + graphicFrame.getId();
  }

  private static ExcelChartSnapshot.Legend snapshotLegend(XSSFChart chart) {
    if (!chart.getCTChart().isSetLegend()) {
      return new ExcelChartSnapshot.Legend.Hidden();
    }
    return new ExcelChartSnapshot.Legend.Visible(
        ExcelChartPoiBridge.fromPoiLegendPosition(
            new org.apache.poi.xddf.usermodel.chart.XDDFChartLegend(chart.getCTChart())
                .getPosition()));
  }

  private static ExcelChartDisplayBlanksAs snapshotDisplayBlanks(XSSFChart chart) {
    return chart.getCTChart().isSetDispBlanksAs()
        ? ExcelChartPoiBridge.fromPoiDisplayBlanks(chart.getCTChart().getDispBlanksAs().getVal())
        : ExcelChartDisplayBlanksAs.GAP;
  }

  private static List<ExcelChartSnapshot.Axis> snapshotAxes(XSSFChart chart) {
    List<ExcelChartSnapshot.Axis> axes = new ArrayList<>();
    for (org.apache.poi.xddf.usermodel.chart.XDDFChartAxis axis : chart.getAxes()) {
      axes.add(
          new ExcelChartSnapshot.Axis(
              ExcelChartPoiBridge.axisKind(axis),
              ExcelChartPoiBridge.fromPoiAxisPosition(axis.getPosition()),
              ExcelChartPoiBridge.fromPoiAxisCrosses(axis.getCrosses()),
              axis.isVisible()));
    }
    return List.copyOf(axes);
  }

  private static List<ExcelChartSnapshot.Series> snapshotSeries(
      XSSFSheet contextSheet, org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series value = data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              ((org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) value)
                  .getCTBarSer()
                  .getTx()));
    }
    return List.copyOf(series);
  }

  private static List<ExcelChartSnapshot.Series> snapshotSeries(
      XSSFSheet contextSheet, org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series value = data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) value)
                  .getCTLineSer()
                  .getTx()));
    }
    return List.copyOf(series);
  }

  private static List<ExcelChartSnapshot.Series> snapshotSeries(
      XSSFSheet contextSheet, org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data) {
    List<ExcelChartSnapshot.Series> series = new ArrayList<>();
    for (int index = 0; index < data.getSeriesCount(); index++) {
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series value = data.getSeries(index);
      series.add(
          snapshotSeries(
              contextSheet,
              value,
              ((org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) value)
                  .getCTPieSer()
                  .getTx()));
    }
    return List.copyOf(series);
  }

  private static ExcelChartSnapshot.Series snapshotSeries(
      XSSFSheet contextSheet,
      org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series series,
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title) {
    return new ExcelChartSnapshot.Series(
        snapshotSeriesTitle(title),
        snapshotDataSource(contextSheet, series.getCategoryData()),
        snapshotDataSource(contextSheet, series.getValuesData()));
  }

  private static ExcelChartSnapshot.DataSource snapshotDataSource(
      XSSFSheet contextSheet, org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    if (source == null) {
      throw new IllegalStateException("Chart series is missing its data source");
    }
    if (source.isReference()) {
      String referenceFormula = source.getDataRangeReference();
      List<String> values = resolvedOrCachedReferenceValues(contextSheet, referenceFormula, source);
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
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    if (referenceFormula != null && !referenceFormula.isBlank()) {
      try {
        return ExcelChartSourceSupport.resolveChartSource(contextSheet, referenceFormula)
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
    return values;
  }

  private static List<String> chartPlotTypeTokens(XSSFChart chart) {
    return chartPlotTypeTokens(chart.getChartSeries());
  }

  private static List<String> chartPlotTypeTokens(
      List<org.apache.poi.xddf.usermodel.chart.XDDFChartData> chartData) {
    List<String> tokens = new ArrayList<>();
    for (org.apache.poi.xddf.usermodel.chart.XDDFChartData value : chartData) {
      tokens.add(ExcelChartPoiBridge.plotTypeToken(value));
    }
    return List.copyOf(tokens);
  }

  private static String titleSummary(ExcelChartSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelChartSnapshot.Bar bar -> titleSummary(bar.title());
      case ExcelChartSnapshot.Line line -> titleSummary(line.title());
      case ExcelChartSnapshot.Pie pie -> titleSummary(pie.title());
      case ExcelChartSnapshot.Unsupported _ -> "";
    };
  }
}
