package dev.erst.gridgrind.core.plan;

import java.util.Objects;

/** One executable mutation with materialized operands and a concrete target. */
public record ResolvedStep(ResolvedTarget target, ResolvedMutationAction action) {
  public ResolvedStep {
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(action, "action must not be null");
    boolean accepted =
        java.util.Arrays.stream(MutationAction.allowedTargetTypesForType(action.authoredType()))
            .anyMatch(type -> type.isInstance(target));
    if (action instanceof ResolvedMutationAction.SetTable table
        && target instanceof TableSelector.ByNameOnSheet selected
        && (!selected.name().equals(table.table().name())
            || !selected.sheetName().equals(table.table().sheetName()))) {
      throw new IllegalArgumentException(
          "SET_TABLE target must match table.name and table.sheetName");
    }
    if (action instanceof StructuredMutationAction.SetPivotTable pivot
        && target instanceof PivotTableSelector.ByNameOnSheet selected
        && (!selected.name().equals(pivot.pivotTable().name())
            || !selected.sheetName().equals(pivot.pivotTable().sheetName()))) {
      throw new IllegalArgumentException(
          "SET_PIVOT_TABLE target must match pivotTable.name and pivotTable.sheetName");
    }
    if (action instanceof StructuredMutationAction.SetNamedRange named
        && target instanceof NamedRangeSelector.ScopedExact selected) {
      boolean matches =
          switch (selected) {
            case NamedRangeSelector.WorkbookScope workbook ->
                workbook.name().equals(named.name())
                    && named.scope()
                        instanceof dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook;
            case NamedRangeSelector.SheetScope sheet ->
                sheet.name().equals(named.name())
                    && named.scope()
                        instanceof dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet scope
                    && scope.sheetName().equals(sheet.sheetName());
          };
      if (!matches)
        throw new IllegalArgumentException(
            "SET_NAMED_RANGE target must match action name and scope");
    }
    if (!accepted)
      throw new IllegalArgumentException(
          "Mutation target does not match " + action.authoredType().getSimpleName());
  }
}
