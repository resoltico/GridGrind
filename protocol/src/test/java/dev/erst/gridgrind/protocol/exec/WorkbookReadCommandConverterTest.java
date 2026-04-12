package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.excel.ExcelPivotTableSelection;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.protocol.dto.PivotTableSelection;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused coverage for read-command converter chart branches. */
class WorkbookReadCommandConverterTest {
  @Test
  void convertsChartReadOperations() {
    WorkbookReadOperation operation = new WorkbookReadOperation.GetCharts("charts", "Budget");
    WorkbookReadCommand.GetCharts command =
        assertInstanceOf(
            WorkbookReadCommand.GetCharts.class,
            WorkbookReadCommandConverter.toReadCommand(operation));

    assertEquals("charts", command.requestId());
    assertEquals("Budget", command.sheetName());
  }

  @Test
  void convertsPivotReadOperations() {
    WorkbookReadOperation.GetPivotTables getPivotTables =
        new WorkbookReadOperation.GetPivotTables(
            "pivots", new PivotTableSelection.ByNames(List.of("Sales Pivot 2026")));
    WorkbookReadOperation.AnalyzePivotTableHealth analyzePivotTableHealth =
        new WorkbookReadOperation.AnalyzePivotTableHealth(
            "pivot-health", new PivotTableSelection.All());

    WorkbookReadCommand.GetPivotTables pivotCommand =
        assertInstanceOf(
            WorkbookReadCommand.GetPivotTables.class,
            WorkbookReadCommandConverter.toReadCommand(getPivotTables));
    WorkbookReadCommand.AnalyzePivotTableHealth pivotHealthCommand =
        assertInstanceOf(
            WorkbookReadCommand.AnalyzePivotTableHealth.class,
            WorkbookReadCommandConverter.toReadCommand(analyzePivotTableHealth));

    assertEquals("pivots", pivotCommand.requestId());
    assertEquals(
        List.of("Sales Pivot 2026"),
        ((ExcelPivotTableSelection.ByNames) pivotCommand.selection()).names());
    assertInstanceOf(ExcelPivotTableSelection.All.class, pivotHealthCommand.selection());
  }
}
