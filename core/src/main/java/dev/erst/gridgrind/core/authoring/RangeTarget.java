package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/** Exact A1 range fluent target. */
public final class RangeTarget {
  private final RangeSelector.ByRange selector;

  RangeTarget(RangeSelector.ByRange selector) {
    this.selector = Objects.requireNonNull(selector, "selector must not be null");
  }

  RangeSelector.ByRange selector() {
    return selector;
  }

  /** Returns one set-range mutation step. */
  public PlannedMutation setRows(List<List<Values.CellValue>> rows) {
    Objects.requireNonNull(rows, "rows must not be null");
    return new PlannedMutation(
        selector,
        new CellMutationAction.SetRange(
            new CellGridSpec.Typed(
                rows.stream()
                    .map(
                        row ->
                            row.stream()
                                .map(Values::toCellInput)
                                .collect(Collectors.toUnmodifiableList()))
                    .collect(Collectors.toUnmodifiableList()))));
  }

  /** Returns one clear-range mutation step. */
  public PlannedMutation clear() {
    return new PlannedMutation(selector, new CellMutationAction.ClearRange());
  }

  /** Returns one merge-cells mutation step. */
  public PlannedMutation merge() {
    return new PlannedMutation(selector, new WorkbookMutationAction.MergeCells());
  }

  /** Returns one unmerge-cells mutation step. */
  public PlannedMutation unmerge() {
    return new PlannedMutation(selector, new WorkbookMutationAction.UnmergeCells());
  }

  /** Returns one data-validations inspection step. */
  public PlannedInspection dataValidations() {
    return new PlannedInspection(selector, SheetQueries.dataValidations());
  }

  /** Returns one conditional-formatting inspection step. */
  public PlannedInspection conditionalFormatting() {
    return new PlannedInspection(selector, SheetQueries.conditionalFormatting());
  }
}
