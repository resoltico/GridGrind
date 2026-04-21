package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Facade over the split chart source, snapshot, and mutation support helpers. */
final class ExcelDrawingChartSupport {
  private ExcelDrawingChartSupport() {}

  static ExcelDrawingObjectSnapshot.Chart snapshotChartDrawingObject(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return ExcelChartSnapshotSupport.snapshotChartDrawingObject(chart, graphicFrame);
  }

  static ExcelChartSnapshot snapshotChart(
      XSSFChart chart, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return ExcelChartSnapshotSupport.snapshotChart(chart, graphicFrame);
  }

  static PreparedChartDefinition prepareChartDefinition(
      XSSFSheet sheet, ExcelChartDefinition definition) {
    return ExcelChartMutationSupport.prepareChartDefinition(sheet, definition);
  }

  static void createChart(XSSFSheet sheet, PreparedChartDefinition definition) {
    ExcelChartMutationSupport.createChart(sheet, definition);
  }

  static void mutateChart(
      XSSFSheet sheet,
      ExcelDrawingController.LocatedShape located,
      XSSFChart chart,
      PreparedChartDefinition definition) {
    ExcelChartMutationSupport.mutateChart(sheet, located, chart, definition);
  }

  static Name resolveDefinedNameReference(XSSFSheet contextSheet, String formula) {
    return ExcelChartSourceSupport.resolveDefinedNameReference(contextSheet, formula);
  }

  static String normalizeAreaFormulaForPoi(String formula) {
    return ExcelChartSourceSupport.normalizeAreaFormulaForPoi(formula);
  }

  static ExcelDrawingController.CellScalar scalarFromFormula(Cell cell) {
    return ExcelChartSourceSupport.scalarFromFormula(cell);
  }

  static String requiredDefinedNameFormula(Name definedName) {
    return ExcelChartSourceSupport.requiredDefinedNameFormula(definedName);
  }

  static XSSFChart chartForGraphicFrame(
      XSSFDrawing drawing, org.apache.poi.xssf.usermodel.XSSFGraphicFrame graphicFrame) {
    return ExcelChartSnapshotSupport.chartForGraphicFrame(drawing, graphicFrame);
  }

  static ExcelChartSnapshot.Title snapshotTitle(XSSFChart chart) {
    return ExcelChartSnapshotSupport.snapshotTitle(chart);
  }

  static String cachedTitleText(XSSFChart chart, String formula) {
    return ExcelChartSnapshotSupport.cachedTitleText(chart, formula);
  }

  static String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    return ExcelChartSnapshotSupport.resolvedTitleFormulaText(chart, formula);
  }

  static boolean barVaryColors(XSSFChart chart) {
    return ExcelChartSnapshotSupport.barVaryColors(chart);
  }

  static boolean lineVaryColors(XSSFChart chart) {
    return ExcelChartSnapshotSupport.lineVaryColors(chart);
  }

  static boolean pieVaryColors(XSSFChart chart) {
    return ExcelChartSnapshotSupport.pieVaryColors(chart);
  }

  static PreparedSeriesTitle prepareSeriesTitle(XSSFSheet sheet, ExcelChartDefinition.Title title) {
    return ExcelChartMutationSupport.prepareSeriesTitle(sheet, title);
  }

  static void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series series,
      PreparedSeriesTitle title) {
    ExcelChartMutationSupport.applySeriesTitle(series, title);
  }

  static void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series series,
      PreparedSeriesTitle title) {
    ExcelChartMutationSupport.applySeriesTitle(series, title);
  }

  static void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series series,
      PreparedSeriesTitle title) {
    ExcelChartMutationSupport.applySeriesTitle(series, title);
  }

  static ExcelChartSnapshot.Title snapshotSeriesTitle(
      org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx title) {
    return ExcelChartSnapshotSupport.snapshotSeriesTitle(title);
  }

  static String titleSummary(ExcelChartSnapshot.Title title) {
    return ExcelChartSnapshotSupport.titleSummary(title);
  }
}
