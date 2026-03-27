package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for workbook-read result invariants. */
class WorkbookReadResultTest {
  @Test
  void workbookSummaryResultRejectsBlankRequestId() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookReadResult.WorkbookSummaryResult(
                    " ", new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 0, false)));

    assertEquals("requestId must not be blank", exception.getMessage());
  }

  @Test
  void cellsResultCopiesCells() {
    WorkbookReadResult.CellsResult result =
        new WorkbookReadResult.CellsResult(
            "cells",
            "Budget",
            List.of(
                new GridGrindResponse.CellReport.BlankReport(
                    "A1", "BLANK", "", defaultStyle(), null, null)));

    assertEquals("A1", result.cells().getFirst().address());
  }

  private static GridGrindResponse.CellStyleReport defaultStyle() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        false,
        false,
        false,
        dev.erst.gridgrind.excel.ExcelHorizontalAlignment.GENERAL,
        dev.erst.gridgrind.excel.ExcelVerticalAlignment.BOTTOM,
        "Aptos",
        new FontHeightReport(220, java.math.BigDecimal.valueOf(11)),
        null,
        false,
        false,
        null,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE);
  }
}
