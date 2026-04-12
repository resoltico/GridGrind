package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
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
}
