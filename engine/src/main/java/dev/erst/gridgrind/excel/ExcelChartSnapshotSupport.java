package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchorSupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;

/** Chart snapshot and chart-readback helpers. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelChartSnapshotSupport {
  private ExcelChartSnapshotSupport() {}

  public static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChartDrawingObject(chart, graphicFrame, null);
  }

  public static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart,
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    ExcelChartSnapshot snapshot = snapshotChart(chart, graphicFrame, formulaRuntime);
    boolean supported =
        snapshot.plots().stream().noneMatch(ExcelChartSnapshot.Unsupported.class::isInstance);
    return new ExcelDrawingObjectSnapshot.Chart(
        ExcelChartTitleSnapshotSupport.resolvedChartName(graphicFrame),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
        supported,
        chartPlotTypeTokens(chart),
        ExcelChartTitleSnapshotSupport.titleSummary(snapshot.title()));
  }

  public static ExcelChartSnapshot snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChart(chart, graphicFrame, null);
  }

  public static ExcelChartSnapshot snapshotChart(
      XSSFChart chart,
      XSSFGraphicFrame graphicFrame,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    List<XDDFChartData> chartData = chart.getChartSeries();
    List<ExcelChartSnapshot.Plot> plots =
        ExcelChartPlotSnapshotSupport.snapshotPlots(chart, graphicFrame, chartData, formulaRuntime);
    return new ExcelChartSnapshot(
        ExcelChartTitleSnapshotSupport.resolvedChartName(graphicFrame),
        ExcelDrawingAnchorSupport.snapshotAnchor(ExcelDrawingAnchorSupport.shapeXml(graphicFrame)),
        ExcelChartTitleSnapshotSupport.snapshotTitle(chart, graphicFrame, formulaRuntime),
        snapshotLegend(chart),
        snapshotDisplayBlanks(chart),
        chart.isPlotOnlyVisibleCells(),
        plots);
  }

  public static boolean barVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfBarChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getBarChartArray(0).getVaryColors());
  }

  public static boolean lineVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfLineChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getLineChartArray(0).getVaryColors());
  }

  public static boolean pieVaryColors(XSSFChart chart) {
    return chart.getCTChart().getPlotArea().sizeOfPieChartArray() > 0
        && truthy(chart.getCTChart().getPlotArea().getPieChartArray(0).getVaryColors());
  }

  static List<String> resolvedOrCachedReferenceValues(
      XSSFSheet contextSheet,
      String referenceFormula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source) {
    return resolvedOrCachedReferenceValues(contextSheet, referenceFormula, source, null);
  }

  static List<String> resolvedOrCachedReferenceValues(
      XSSFSheet contextSheet,
      String referenceFormula,
      org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> source,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartSeriesSnapshotSupport.resolvedOrCachedReferenceValues(
        contextSheet, referenceFormula, source, formulaRuntime);
  }

  private static List<String> chartPlotTypeTokens(XSSFChart chart) {
    return chartPlotTypeTokens(chart.getChartSeries());
  }

  private static List<String> chartPlotTypeTokens(List<XDDFChartData> chartData) {
    List<String> tokens = new ArrayList<>();
    for (XDDFChartData value : chartData) {
      tokens.add(ExcelChartDataFamilyPoiBridge.plotTypeToken(value));
    }
    return List.copyOf(tokens);
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

  private static boolean truthy(CTBoolean value) {
    return value != null && value.getVal();
  }
}
