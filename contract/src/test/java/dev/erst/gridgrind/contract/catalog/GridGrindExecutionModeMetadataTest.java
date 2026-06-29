package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for execution-mode metadata branches that are public contract helpers. */
class GridGrindExecutionModeMetadataTest {
  @Test
  void fullXssfExposesItsCanonicalModeIdAndSummary() {
    GridGrindExecutionModeMetadata.FullXssfMode mode = GridGrindExecutionModeMetadata.fullXssf();

    assertEquals("FULL_XSSF", mode.modeId());
    assertTrue(mode.catalogSummary().contains("full workbook read and write"));
  }

  @Test
  void eventReadRejectsAnEmptyAllowedQueryList() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindExecutionModeMetadata.EventReadMode(
                    ExecutionModeInput.EventRead.class,
                    List.of(),
                    CalculationStrategyInput.DoNotCalculate.class,
                    false));

    assertEquals("allowedQueries must not be empty", failure.getMessage());
  }

  @Test
  void streamingWriteRejectsAnEmptyAllowedActionList() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindExecutionModeMetadata.StreamingWriteMode(
                    ExecutionModeInput.StreamingWrite.class,
                    WorkbookPlan.WorkbookSource.New.class,
                    List.of(),
                    CalculationStrategyInput.DoNotCalculate.class,
                    true));

    assertEquals("allowedActions must not be empty", failure.getMessage());
  }
}
