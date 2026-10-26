package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.util.List;
import java.util.Optional;

/** Shared chart-authoring and chart-readback helpers for engine tests. */
final class ExcelChartTestSupport {
  private ExcelChartTestSupport() {}

  static DrawingAnchorSpec.TwoCell anchor(int fromColumn, int fromRow, int toColumn, int toRow) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(fromColumn, fromRow, 0, 0),
        new DrawingMarkerSpec(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  static ExcelChartDefinition.DataSource.Reference ref(String formula) {
    return new ExcelChartDefinition.DataSource.Reference(formula);
  }

  static ExcelChartDefinition lineChart(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
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
      DrawingAnchorSpec.TwoCell anchor,
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
                Optional.empty(),
                Optional.empty(),
                axes(),
                series)));
  }

  static ExcelChartDefinition pieChart(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
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
        List.of(
            new ExcelChartDefinition.Pie(
                varyColors, Optional.ofNullable(firstSliceAngle), series)));
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
    sheet.cells().setCell("A1", new CellValue.Text("Month"));
    sheet.cells().setCell("B1", new CellValue.Text("Plan"));
    sheet.cells().setCell("C1", new CellValue.Text("Actual"));
    sheet.cells().setCell("A2", new CellValue.Text("Jan"));
    sheet.cells().setCell("B2", new CellValue.NumberValue(10d));
    sheet.cells().setCell("C2", new CellValue.NumberValue(12d));
    sheet.cells().setCell("A3", new CellValue.Text("Feb"));
    sheet.cells().setCell("B3", new CellValue.NumberValue(18d));
    sheet.cells().setCell("C3", new CellValue.NumberValue(16d));
    sheet.cells().setCell("A4", new CellValue.Text("Mar"));
    sheet.cells().setCell("B4", new CellValue.NumberValue(15d));
    sheet.cells().setCell("C4", new CellValue.NumberValue(21d));
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
        .names()
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartCategories",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range(sheetName, "A2:A4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartPlan",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range(sheetName, "B2:B4")))
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "ChartActual",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range(sheetName, "C2:C4")));
  }
}
