package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for defined-name target normalization into typed GridGrind range targets. */
class ExcelNamedRangeTargetsTest {
  @Test
  void resolvesWorkbookScopedRangeTargetsWithRepeatedSheetPrefixes() {
    Optional<ExcelNamedRangeTarget> target =
        ExcelNamedRangeTargets.resolveTarget(
            "Budget!$A$1:Budget!$B$4", new ExcelNamedRangeScope.WorkbookScope());

    assertEquals(Optional.of(new ExcelNamedRangeTarget("Budget", "A1:B4")), target);
  }

  @Test
  void resolvesSheetScopedRangeTargetsWithRepeatedSheetPrefixes() {
    Optional<ExcelNamedRangeTarget> target =
        ExcelNamedRangeTargets.resolveTarget(
            "Budget!$A$1:Budget!$B$4", new ExcelNamedRangeScope.SheetScope("Budget"));

    assertEquals(Optional.of(new ExcelNamedRangeTarget("Budget", "A1:B4")), target);
  }

  @Test
  void resolvesSingleCellAndSingleRowRangeTargets() {
    assertEquals(
        Optional.of(new ExcelNamedRangeTarget("Budget", "A1")),
        ExcelNamedRangeTargets.resolveTarget(
            "Budget!$A$1", new ExcelNamedRangeScope.WorkbookScope()));
    assertEquals(
        Optional.of(new ExcelNamedRangeTarget("Budget", "A1:B1")),
        ExcelNamedRangeTargets.resolveTarget(
            "Budget!$A$1:Budget!$B$1", new ExcelNamedRangeScope.WorkbookScope()));
  }

  @Test
  void returnsEmptyForNonRangeDefinedNameFormulas() {
    Optional<ExcelNamedRangeTarget> target =
        ExcelNamedRangeTargets.resolveTarget(
            "SUM(Budget!$A$1:Budget!$B$4)", new ExcelNamedRangeScope.WorkbookScope());

    assertTrue(target.isEmpty());
  }

  @Test
  void returnsEmptyForBlankInvalidAndWorkbookScopedSheetlessRanges() {
    assertThrows(
        NullPointerException.class,
        () -> ExcelNamedRangeTargets.resolveTarget(null, new ExcelNamedRangeScope.WorkbookScope()));
    assertTrue(
        ExcelNamedRangeTargets.resolveTarget(" ", new ExcelNamedRangeScope.WorkbookScope())
            .isEmpty());
    assertTrue(
        ExcelNamedRangeTargets.resolveTarget(
                "Budget!$A$1,Budget!$B$4", new ExcelNamedRangeScope.WorkbookScope())
            .isEmpty());
    assertTrue(
        ExcelNamedRangeTargets.resolveTarget(
                "Budget!$A$1,", new ExcelNamedRangeScope.WorkbookScope())
            .isEmpty());
    assertTrue(
        ExcelNamedRangeTargets.resolveTarget("A1:B4", new ExcelNamedRangeScope.WorkbookScope())
            .isEmpty());
  }

  @Test
  void resolvesSheetScopedSheetlessRangesAndRejectsMismatchedSheetPrefixes() {
    assertEquals(
        Optional.of(new ExcelNamedRangeTarget("Budget", "A1:B4")),
        ExcelNamedRangeTargets.resolveTarget(
            "A1:B4", new ExcelNamedRangeScope.SheetScope("Budget")));
    assertTrue(
        ExcelNamedRangeTargets.resolveTarget(
                "Budget!$A$1:Summary!$B$4", new ExcelNamedRangeScope.WorkbookScope())
            .isEmpty());
  }
}
