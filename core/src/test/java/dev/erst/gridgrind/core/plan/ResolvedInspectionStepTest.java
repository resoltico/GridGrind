package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import org.junit.jupiter.api.Test;

/** Executable inspections retain canonical queries and enforce their target contract. */
class ResolvedInspectionStepTest {
  @Test
  void preservesCanonicalQueryAndRejectsWrongTargetKind() {
    var query = new SheetIntrospectionQuery.GetCells();
    var target = new CellSelector.ByAddress(new SheetName("Grid"), new CellAddress("B2"));
    var step = new ResolvedInspectionStep(new StepId("read"), target, query);
    assertSame(query, step.query());
    assertSame(target, step.target());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedInspectionStep(new StepId("bad"), new WorkbookSelector.Current(), query));
  }
}
