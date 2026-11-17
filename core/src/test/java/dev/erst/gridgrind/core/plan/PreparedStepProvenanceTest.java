package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.ResolvedCellText;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Preparation binds external leaves without rewriting authored identity or unrelated operands. */
class PreparedStepProvenanceTest {
  private final StepId id = new StepId("same");
  private final CellSelector.ByAddress cell =
      new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"));

  @Test
  void retainsOriginalMutationAndDerivesOnlyItsBoundValues() {
    var authored =
        new MutationStep(
            id,
            cell,
            new CellMutationAction.SetCell(
                new CellValueSpec.Text(TextSourceSpec.utf8File("value.txt"))));
    var bound =
        new MutationStep(
            id,
            cell,
            new CellMutationAction.SetCell(new CellValueSpec.Text(TextSourceSpec.inline("Bound"))));
    var prepared =
        assertInstanceOf(PreparedMutationStep.class, PreparedStep.prepare(authored, bound));
    assertSame(authored, prepared.authoredStep());
    assertSame(cell, prepared.target());
    assertEquals(
        new ResolvedCellText("Bound"),
        assertInstanceOf(ResolvedMutationAction.SetCell.class, prepared.action()).value());
  }

  @Test
  void rejectsChangedActionValueOperationSelectorAndQuerySharingTheSameId() {
    var original =
        new MutationStep(
            id, cell, new CellMutationAction.SetCell(new CellValueSpec.NumberValue(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            PreparedStep.prepare(
                original,
                new MutationStep(
                    id, cell, new CellMutationAction.SetCell(new CellValueSpec.NumberValue(2)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            PreparedStep.prepare(
                original, new MutationStep(id, cell, new CellMutationAction.ClearComment())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            PreparedStep.prepare(
                original,
                new MutationStep(
                    id,
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B1")),
                    original.action())));
    var inspection = new InspectionStep(id, cell, new SheetIntrospectionQuery.GetCells());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            PreparedStep.prepare(
                inspection,
                new InspectionStep(id, cell, new SheetIntrospectionQuery.GetSheetSummary())));
    var assertion =
        new AssertionStep(id, cell, new CellAssertion.Value(new CellValue.NumberValue(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            PreparedStep.prepare(
                assertion,
                new AssertionStep(
                    id, cell, new CellAssertion.Value(new CellValue.NumberValue(2)))));
  }

  @Test
  void rejectsUnboundExternalLeavesAndChangedInlineText() {
    var external =
        new AssertionStep(
            id,
            key(TextSourceSpec.utf8File("key.txt")),
            new CellAssertion.Value(new CellValue.NumberValue(1)));
    assertThrows(IllegalArgumentException.class, () -> PreparedStep.prepare(external));
    assertThrows(
        IllegalArgumentException.class,
        () -> new PreparedAssertionStep(external, external.target()));
    var inline =
        new AssertionStep(id, key(TextSourceSpec.inline("Original")), external.assertion());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            PreparedStep.prepare(
                inline,
                new AssertionStep(
                    id, key(TextSourceSpec.inline("Unrelated")), inline.assertion())));
  }

  @Test
  void resolvedPlanRejectsAnUnrelatedOriginalEvenIfIdAndKindAgree() {
    var authored =
        new AssertionStep(id, cell, new CellAssertion.Value(new CellValue.NumberValue(1)));
    var unrelated =
        new AssertionStep(id, cell, new CellAssertion.Value(new CellValue.NumberValue(2)));
    var plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(authored));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedPlan(plan, List.of(PreparedStep.prepare(unrelated))));
  }

  private static TableCellSelector.ByColumnName key(TextSourceSpec value) {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("Orders")),
            "Item",
            new CellValueSpec.Text(value)),
        "Amount");
  }
}
