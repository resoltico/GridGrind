package dev.erst.gridgrind.excel.drawing;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.excel.ExcelChartMutationSupport;
import dev.erst.gridgrind.excel.ExcelChartRelationSupport;
import dev.erst.gridgrind.excel.ExcelChartSnapshotSupport;
import dev.erst.gridgrind.excel.ExcelChartTitleSnapshotSupport;
import dev.erst.gridgrind.excel.ExcelFormulaRuntime;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;

/** Facade over the split chart source, snapshot, and mutation support helpers. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelDrawingChartSupport {
  private ExcelDrawingChartSupport() {}

  public static DrawingObjectReport.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChartDrawingObject(chart, graphicFrame, null);
  }

  public static DrawingObjectReport.Chart snapshotChartDrawingObject(
      XSSFChart chart,
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartSnapshotSupport.snapshotChartDrawingObject(
        chart, graphicFrame, formulaRuntime);
  }

  public static ChartReport snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChart(chart, graphicFrame, null);
  }

  public static ChartReport snapshotChart(
      XSSFChart chart,
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartSnapshotSupport.snapshotChart(chart, graphicFrame, formulaRuntime);
  }

  public static void createChart(XSSFSheet sheet, ResolvedChart definition) {
    createChart(sheet, definition, null);
  }

  public static void createChart(
      XSSFSheet sheet, ResolvedChart definition, @Nullable ExcelFormulaRuntime formulaRuntime) {
    ExcelChartMutationSupport.createChart(sheet, definition, formulaRuntime);
  }

  public static void validateChart(XSSFSheet sheet, ResolvedChart definition) {
    validateChart(sheet, definition, null);
  }

  public static void validateChart(
      XSSFSheet sheet, ResolvedChart definition, @Nullable ExcelFormulaRuntime formulaRuntime) {
    ExcelChartMutationSupport.validateChart(sheet, definition, formulaRuntime);
  }

  public static @Nullable XSSFChart chartForGraphicFrame(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return ExcelChartRelationSupport.chartForGraphicFrame(drawing, graphicFrame);
  }

  public static ChartReport.Title snapshotTitle(XSSFChart chart) {
    return ExcelChartTitleSnapshotSupport.snapshotTitle(chart);
  }

  public static String cachedTitleText(XSSFChart chart, String formula) {
    return ExcelChartTitleSnapshotSupport.cachedTitleText(chart, formula);
  }

  public static String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    return ExcelChartTitleSnapshotSupport.resolvedTitleFormulaText(chart, formula);
  }

  public static boolean barVaryColors(XSSFChart chart) {
    return ExcelChartSnapshotSupport.barVaryColors(chart);
  }

  public static boolean lineVaryColors(XSSFChart chart) {
    return ExcelChartSnapshotSupport.lineVaryColors(chart);
  }

  public static boolean pieVaryColors(XSSFChart chart) {
    return ExcelChartSnapshotSupport.pieVaryColors(chart);
  }

  public static ChartReport.Title snapshotSeriesTitle(
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title) {
    return ExcelChartTitleSnapshotSupport.snapshotSeriesTitle(title);
  }

  public static ChartReport.Title snapshotSeriesTitle(
      XSSFSheet contextSheet,
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartTitleSnapshotSupport.snapshotSeriesTitle(contextSheet, title, formulaRuntime);
  }

  public static String titleSummary(ChartReport.Title title) {
    return ExcelChartTitleSnapshotSupport.titleSummary(title);
  }
}
