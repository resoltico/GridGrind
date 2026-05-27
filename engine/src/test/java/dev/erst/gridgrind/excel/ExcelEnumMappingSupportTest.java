package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import java.util.Map;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.junit.jupiter.api.Test;

/** Regression coverage for exact enum-mapping helper failure modes. */
class ExcelEnumMappingSupportTest {
  @Test
  void exactEnumMapRejectsIncompleteCoverage() {
    IllegalStateException exception =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelEnumMappingSupport.exactEnumMap(
                    ExcelChartPlotType.class,
                    "chart type mapping",
                    Map.of(ExcelChartPlotType.AREA, ChartTypes.AREA)));

    assertTrue(exception.getMessage().contains("must cover every ExcelChartPlotType constant"));
  }

  @Test
  void reverseExactEnumMapRejectsDuplicateTargets() {
    IllegalStateException exception =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelEnumMappingSupport.reverseExactEnumMap(
                    ExcelFillPattern.class,
                    "fill pattern mapping",
                    Map.of(
                        FillPatternType.NO_FILL,
                        ExcelFillPattern.NONE,
                        FillPatternType.SOLID_FOREGROUND,
                        ExcelFillPattern.NONE)));

    assertTrue(exception.getMessage().contains("to the same target"));
  }

  @Test
  void reverseExactEnumMapRejectsIncompleteReverseCoverage() {
    IllegalStateException exception =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelEnumMappingSupport.reverseExactEnumMap(
                    ExcelFillPattern.class,
                    "fill pattern mapping",
                    Map.of(FillPatternType.NO_FILL, ExcelFillPattern.NONE)));

    assertTrue(exception.getMessage().contains("must cover every ExcelFillPattern constant"));
  }

  @Test
  void requireMappedValueRejectsUnknownKeys() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelEnumMappingSupport.requireMappedValue(
                    Map.of("known", 1), "missing", "test key"));

    assertEquals("Unsupported test key: missing", exception.getMessage());
  }
}
