package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.engine.names.NamedRangeAdapter;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for defined-name target normalization into typed GridGrind range targets. */
class NamedRangeTargetResolutionTest {
  @Test
  void resolvesWorkbookScopedRangeTargetsWithRepeatedSheetPrefixes() {
    Optional<NamedRangeTarget> target =
        NamedRangeAdapter.resolveTarget(
            "Budget!$A$1:Budget!$B$4", new NamedRangeScopeReport.Workbook());

    assertEquals(
        Optional.of(
            NamedRangeTarget.range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B4"))),
        target);
  }

  @Test
  void resolvesSheetScopedRangeTargetsWithRepeatedSheetPrefixes() {
    Optional<NamedRangeTarget> target =
        NamedRangeAdapter.resolveTarget(
            "Budget!$A$1:Budget!$B$4",
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")));

    assertEquals(
        Optional.of(
            NamedRangeTarget.range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B4"))),
        target);
  }

  @Test
  void resolvesSingleCellAndSingleRowRangeTargets() {
    assertEquals(
        Optional.of(
            NamedRangeTarget.range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"))),
        NamedRangeAdapter.resolveTarget("Budget!$A$1", new NamedRangeScopeReport.Workbook()));
    assertEquals(
        Optional.of(
            NamedRangeTarget.range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B1"))),
        NamedRangeAdapter.resolveTarget(
            "Budget!$A$1:Budget!$B$1", new NamedRangeScopeReport.Workbook()));
  }

  @Test
  void returnsEmptyForNonRangeDefinedNameFormulas() {
    Optional<NamedRangeTarget> target =
        NamedRangeAdapter.resolveTarget(
            "SUM(Budget!$A$1:Budget!$B$4)", new NamedRangeScopeReport.Workbook());

    assertTrue(target.isEmpty());
  }

  @Test
  void returnsEmptyForBlankInvalidAndWorkbookScopedSheetlessRanges() {
    assertThrows(
        NullPointerException.class,
        () -> NamedRangeAdapter.resolveTarget(null, new NamedRangeScopeReport.Workbook()));
    assertTrue(
        NamedRangeAdapter.resolveTarget(" ", new NamedRangeScopeReport.Workbook()).isEmpty());
    assertTrue(
        NamedRangeAdapter.resolveTarget(
                "Budget!$A$1,Budget!$B$4", new NamedRangeScopeReport.Workbook())
            .isEmpty());
    assertTrue(
        NamedRangeAdapter.resolveTarget("Budget!$A$1,", new NamedRangeScopeReport.Workbook())
            .isEmpty());
    assertTrue(
        NamedRangeAdapter.resolveTarget("A1:B4", new NamedRangeScopeReport.Workbook()).isEmpty());
  }

  @Test
  void resolvesSheetScopedSheetlessRangesAndRejectsMismatchedSheetPrefixes() {
    assertEquals(
        Optional.of(
            NamedRangeTarget.range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B4"))),
        NamedRangeAdapter.resolveTarget(
            "A1:B4",
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"))));
    assertTrue(
        NamedRangeAdapter.resolveTarget(
                "Budget!$A$1:Summary!$B$4", new NamedRangeScopeReport.Workbook())
            .isEmpty());
  }
}
