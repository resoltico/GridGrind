package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for execution-mode metadata branches that are public contract helpers. */
class BackendCapabilitiesTest {
  @Test
  void fullXssfExposesItsCanonicalModeIdAndSummary() {
    BackendCapabilities.FullXssfMode mode = BackendCapabilities.fullXssf();

    assertEquals("FULL_XSSF", mode.modeId());
    assertTrue(mode.catalogSummary().contains("full workbook read and write"));
  }

  @Test
  void eventReadRejectsAnEmptyAllowedQueryList() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new BackendCapabilities.EventReadMode(
                    ExecutionModeSpec.EventRead.class,
                    List.of(),
                    CalculationStrategySpec.DoNotCalculate.class,
                    false));

    assertEquals("allowedQueries must not be empty", failure.getMessage());
  }

  @Test
  void streamingWriteRejectsAnEmptyAllowedActionList() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new BackendCapabilities.StreamingWriteMode(
                    ExecutionModeSpec.StreamingWrite.class,
                    WorkbookPlan.WorkbookSource.New.class,
                    List.of(),
                    CalculationStrategySpec.DoNotCalculate.class,
                    true));

    assertEquals("allowedActions must not be empty", failure.getMessage());
  }
}
