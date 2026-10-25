package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import java.util.ArrayList;
import java.util.List;

/** Rejects statically provable worksheet work that cannot run safely in full XSSF. */
final class WorkbookStaticMaterializationValidation {
  static final long MAX_MATERIALIZED_WORK_ITEMS = 250_000L; // LIM-039

  private WorkbookStaticMaterializationValidation() {}

  static List<WorkbookStaticViolation> validate(List<WorkbookStaticStep> steps) {
    long totalWorkItems = 0L;
    List<WorkbookStaticViolation> violations = new ArrayList<>();
    for (WorkbookStaticStep staticStep : steps) {
      if (staticStep.value().isEmpty()
          || !(staticStep.value().orElseThrow() instanceof MutationStep mutation)) {
        continue;
      }
      WorkbookStaticWorkEstimate estimate = workEstimate(staticStep.index(), mutation);
      if (estimate instanceof InvalidWorkbookStaticWorkEstimate invalid) {
        violations.add(invalid.violation());
        continue;
      }
      ValidWorkbookStaticWorkEstimate valid = (ValidWorkbookStaticWorkEstimate) estimate;
      totalWorkItems = Math.addExact(totalWorkItems, valid.workItems());
      if (totalWorkItems > MAX_MATERIALIZED_WORK_ITEMS) {
        violations.add(
            new WorkbookStaticViolation(
                valid.jsonPath(),
                "materialized worksheet work must not exceed "
                    + MAX_MATERIALIZED_WORK_ITEMS
                    + " items per plan; this step raises the total to "
                    + totalWorkItems));
      }
    }
    return List.copyOf(violations);
  }

  private static WorkbookStaticWorkEstimate workEstimate(int stepIndex, MutationStep mutation) {
    if (mutation.target() instanceof RangeSelector.ByRange range) {
      if (mutation.action() instanceof CellMutationAction.SetRange setRange
          && authoredCells(setRange) != range.cellCount()) {
        return new InvalidWorkbookStaticWorkEstimate(
            new WorkbookStaticViolation(
                "steps[" + stepIndex + "].action.rows",
                "range dimensions do not match provided values: "
                    + range.range()
                    + " expects "
                    + range.cellCount()
                    + " cells but received "
                    + authoredCells(setRange)));
      }
      return new ValidWorkbookStaticWorkEstimate(
          materializedCells(mutation.action(), range), "steps[" + stepIndex + "].target.range");
    }
    if (mutation.target() instanceof RowBandSelector.Span rows) {
      return new ValidWorkbookStaticWorkEstimate(
          materializedRows(mutation.action(), rows),
          "steps[" + stepIndex + "].target.lastRowIndex");
    }
    return new ValidWorkbookStaticWorkEstimate(0L, "steps[" + stepIndex + "].target");
  }

  private static long materializedCells(
      dev.erst.gridgrind.contract.action.MutationAction action, RangeSelector.ByRange range) {
    return switch (action) {
      case CellMutationAction.ApplyStyle _ -> range.cellCount();
      case CellMutationAction.ClearRange _ -> range.cellCount();
      case CellMutationAction.SetArrayFormula _ -> range.cellCount();
      case CellMutationAction.SetRange setRange -> authoredCells(setRange);
      default -> 0L;
    };
  }

  private static long materializedRows(
      dev.erst.gridgrind.contract.action.MutationAction action, RowBandSelector.Span rows) {
    return switch (action) {
      case WorkbookMutationAction.SetRowHeight _ -> rowCount(rows);
      case WorkbookMutationAction.SetRowVisibility _ -> rowCount(rows);
      case WorkbookMutationAction.GroupRows _ -> rowCount(rows);
      case WorkbookMutationAction.UngroupRows _ -> rowCount(rows);
      default -> 0L;
    };
  }

  private static long authoredCells(CellMutationAction.SetRange setRange) {
    var rows = setRange.rows().toCellInputRows();
    return (long) rows.size() * rows.getFirst().size();
  }

  private static long rowCount(RowBandSelector.Span rows) {
    return (long) rows.lastRowIndex() - rows.firstRowIndex() + 1L;
  }
}
