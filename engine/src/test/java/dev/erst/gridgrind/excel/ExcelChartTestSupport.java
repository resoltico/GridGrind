package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import java.util.List;

/** Shared chart-authoring and chart-readback helpers for engine tests. */
final class ExcelChartTestSupport {
  private ExcelChartTestSupport() {}

  static ExcelDrawingAnchor.TwoCell anchor(int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow),
        new ExcelDrawingMarker(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  static ExcelChartDefinition.DataSource.Reference ref(String formula) {
    return new ExcelChartDefinition.DataSource.Reference(formula);
  }

  static ExcelChartDefinition lineChart(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      ExcelChartDefinition.Title title,
      ExcelChartDefinition.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      List<ExcelChartDefinition.Series> series) {
    return new ExcelChartDefinition(
        name,
        anchor,
        title,
        legend,
        displayBlanksAs,
        plotOnlyVisibleCells,
        List.of(
            new ExcelChartDefinition.Line(
                varyColors, ExcelChartGrouping.STANDARD, axes(), series)));
  }

  static ExcelChartDefinition barChart(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      ExcelChartDefinition.Title title,
      ExcelChartDefinition.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      ExcelChartBarDirection barDirection,
      List<ExcelChartDefinition.Series> series) {
    return new ExcelChartDefinition(
        name,
        anchor,
        title,
        legend,
        displayBlanksAs,
        plotOnlyVisibleCells,
        List.of(
            new ExcelChartDefinition.Bar(
                varyColors,
                barDirection,
                ExcelChartBarGrouping.CLUSTERED,
                null,
                null,
                axes(),
                series)));
  }

  static ExcelChartDefinition pieChart(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      ExcelChartDefinition.Title title,
      ExcelChartDefinition.Legend legend,
      ExcelChartDisplayBlanksAs displayBlanksAs,
      boolean plotOnlyVisibleCells,
      boolean varyColors,
      Integer firstSliceAngle,
      List<ExcelChartDefinition.Series> series) {
    return new ExcelChartDefinition(
        name,
        anchor,
        title,
        legend,
        displayBlanksAs,
        plotOnlyVisibleCells,
        List.of(new ExcelChartDefinition.Pie(varyColors, firstSliceAngle, series)));
  }

  static ExcelChartSnapshot chart(List<ExcelChartSnapshot> charts, String name) {
    return charts.stream()
        .filter(snapshot -> name.equals(snapshot.name()))
        .findFirst()
        .orElseThrow();
  }

  static <T extends ExcelChartSnapshot.Plot> T singlePlot(
      ExcelChartSnapshot chart, Class<T> plotType) {
    if (chart.plots().size() != 1) {
      throw new IllegalArgumentException(
          "Expected one plot for chart '" + chart.name() + "' but found " + chart.plots().size());
    }
    return plotType.cast(chart.plots().getFirst());
  }

  private static List<ExcelChartDefinition.Axis> axes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  static void seedChartData(ExcelSheet sheet) {
    sheet.setCell("A1", ExcelCellValue.text("Month"));
    sheet.setCell("B1", ExcelCellValue.text("Plan"));
    sheet.setCell("C1", ExcelCellValue.text("Actual"));
    sheet.setCell("A2", ExcelCellValue.text("Jan"));
    sheet.setCell("B2", ExcelCellValue.number(10d));
    sheet.setCell("C2", ExcelCellValue.number(12d));
    sheet.setCell("A3", ExcelCellValue.text("Feb"));
    sheet.setCell("B3", ExcelCellValue.number(18d));
    sheet.setCell("C3", ExcelCellValue.number(16d));
    sheet.setCell("A4", ExcelCellValue.text("Mar"));
    sheet.setCell("B4", ExcelCellValue.number(15d));
    sheet.setCell("C4", ExcelCellValue.number(21d));
  }

  static void seedChartData(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Month");
    sheet.getRow(0).createCell(1).setCellValue("Plan");
    sheet.getRow(0).createCell(2).setCellValue("Actual");
    sheet.createRow(1).createCell(0).setCellValue("Jan");
    sheet.getRow(1).createCell(1).setCellValue(10d);
    sheet.getRow(1).createCell(2).setCellValue(12d);
    sheet.createRow(2).createCell(0).setCellValue("Feb");
    sheet.getRow(2).createCell(1).setCellValue(18d);
    sheet.getRow(2).createCell(2).setCellValue(16d);
    sheet.createRow(3).createCell(0).setCellValue("Mar");
    sheet.getRow(3).createCell(1).setCellValue(15d);
    sheet.getRow(3).createCell(2).setCellValue(21d);
  }

  static void seedChartNamedRanges(ExcelWorkbook workbook, String sheetName) {
    workbook
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget(sheetName, "A2:A4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartPlan",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget(sheetName, "B2:B4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartActual",
                new ExcelNamedRangeScope.WorkbookScope(),
                new ExcelNamedRangeTarget(sheetName, "C2:C4")));
  }
}
