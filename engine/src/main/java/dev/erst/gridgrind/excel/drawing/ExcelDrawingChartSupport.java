package dev.erst.gridgrind.excel.drawing;

import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartMutationSupport;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelChartSnapshotSupport;
import dev.erst.gridgrind.excel.ExcelFormulaRuntime;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;

/** Facade over the split chart source, snapshot, and mutation support helpers. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelDrawingChartSupport {
  private ExcelDrawingChartSupport() {}

  public static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChartDrawingObject(chart, graphicFrame, null);
  }

  public static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart,
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartSnapshotSupport.snapshotChartDrawingObject(
        chart, graphicFrame, formulaRuntime);
  }

  public static ExcelChartSnapshot snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return snapshotChart(chart, graphicFrame, null);
  }

  public static ExcelChartSnapshot snapshotChart(
      XSSFChart chart,
      org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartSnapshotSupport.snapshotChart(chart, graphicFrame, formulaRuntime);
  }

  public static void createChart(XSSFSheet sheet, ExcelChartDefinition definition) {
    createChart(sheet, definition, null);
  }

  public static void createChart(
      XSSFSheet sheet,
      ExcelChartDefinition definition,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    ExcelChartMutationSupport.createChart(sheet, definition, formulaRuntime);
  }

  public static void validateChart(XSSFSheet sheet, ExcelChartDefinition definition) {
    validateChart(sheet, definition, null);
  }

  public static void validateChart(
      XSSFSheet sheet,
      ExcelChartDefinition definition,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    ExcelChartMutationSupport.validateChart(sheet, definition, formulaRuntime);
  }

  public static @Nullable XSSFChart chartForGraphicFrame(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return ExcelChartSnapshotSupport.chartForGraphicFrame(drawing, graphicFrame);
  }

  public static ExcelChartSnapshot.Title snapshotTitle(XSSFChart chart) {
    return ExcelChartSnapshotSupport.snapshotTitle(chart);
  }

  public static String cachedTitleText(XSSFChart chart, String formula) {
    return ExcelChartSnapshotSupport.cachedTitleText(chart, formula);
  }

  public static String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    return ExcelChartSnapshotSupport.resolvedTitleFormulaText(chart, formula);
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

  public static ExcelChartSnapshot.Title snapshotSeriesTitle(
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title) {
    return ExcelChartSnapshotSupport.snapshotSeriesTitle(title);
  }

  public static ExcelChartSnapshot.Title snapshotSeriesTitle(
      XSSFSheet contextSheet,
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    return ExcelChartSnapshotSupport.snapshotSeriesTitle(contextSheet, title, formulaRuntime);
  }

  public static String titleSummary(ExcelChartSnapshot.Title title) {
    return ExcelChartSnapshotSupport.titleSummary(title);
  }
}
