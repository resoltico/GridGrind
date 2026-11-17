package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.Assertion;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import org.jspecify.annotations.Nullable;

/** Resolves semantic selectors into exact executable targets without inventing new semantics. */
final class SemanticSelectorResolver {
  private static final Set<Class<? extends MutationAction>> EXACT_CELL_MUTATION_ACTIONS =
      Set.of(
          CellMutationAction.SetCell.class,
          CellMutationAction.SetHyperlink.class,
          CellMutationAction.ClearHyperlink.class,
          CellMutationAction.SetComment.class,
          CellMutationAction.ClearComment.class);
  private static final Set<Class<? extends InspectionQuery>> EXACT_CELL_INSPECTION_QUERIES =
      Set.of(
          SheetIntrospectionQuery.GetCells.class,
          SheetIntrospectionQuery.GetHyperlinks.class,
          SheetIntrospectionQuery.GetComments.class);
  private static final Set<Class<? extends Assertion>> EXACT_CELL_ASSERTIONS =
      Set.of(
          CellAssertion.Value.class,
          CellAssertion.DisplayValue.class,
          CellAssertion.FormulaText.class,
          CellAssertion.CellStyle.class);

  private final TableCellResolver tableCells;

  SemanticSelectorResolver(WorkbookExecutionEngine workbookEngine) {
    this.tableCells = new TableCellResolver(workbookEngine);
  }

  Selector resolveMutationTarget(XssfWorkbook workbook, Selector target, MutationAction action) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(action, "action must not be null");
    return EXACT_CELL_MUTATION_ACTIONS.contains(action.getClass())
        ? resolveExactCellTarget(workbook, target)
        : target;
  }

  ResolvedInspectionTarget resolveInspectionTarget(
      StepId stepId, XssfWorkbook workbook, Selector target, InspectionQuery query) {
    Objects.requireNonNull(stepId, "stepId must not be null");
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(query, "query must not be null");
    return EXACT_CELL_INSPECTION_QUERIES.contains(query.getClass())
        ? resolveExactCellInspectionTarget(stepId, workbook, target, query)
        : ResolvedInspectionTarget.direct(target);
  }

  Selector resolveAssertionTarget(XssfWorkbook workbook, Selector target, Assertion assertion) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(assertion, "assertion must not be null");
    return EXACT_CELL_ASSERTIONS.contains(assertion.getClass())
        ? resolveExactCellTarget(workbook, target)
        : target;
  }

  private Selector resolveExactCellTarget(XssfWorkbook workbook, Selector target) {
    if (target instanceof TableCellSelector.ByColumnName tableCell) {
      var resolvedCell = tableCells.resolve(workbook, tableCell, false);
      return new CellSelector.ByAddress(
          new SheetName(resolvedCell.sheetName()),
          Objects.requireNonNull(resolvedCell.address(), "resolved cell address must be present"));
    }
    return target;
  }

  private ResolvedInspectionTarget resolveExactCellInspectionTarget(
      StepId stepId, XssfWorkbook workbook, Selector target, InspectionQuery query) {
    if (!(target instanceof TableCellSelector.ByColumnName tableCell)) {
      return ResolvedInspectionTarget.direct(target);
    }
    var resolvedCell = tableCells.resolve(workbook, tableCell, true);
    if (!resolvedCell.matched()) {
      return ResolvedInspectionTarget.shortCircuit(
          emptyInspectionResult(stepId, resolvedCell.sheetName(), query));
    }
    return ResolvedInspectionTarget.direct(
        new CellSelector.ByAddress(
            new SheetName(resolvedCell.sheetName()),
            Objects.requireNonNull(
                resolvedCell.address(), "resolved cell address must be present")));
  }

  private static InspectionResult emptyInspectionResult(
      StepId stepId, String sheetName, InspectionQuery query) {
    if (query instanceof SheetIntrospectionQuery.GetCells) {
      return new SheetInspectionResult.CellsResult(stepId, new SheetLabel(sheetName), List.of());
    }
    if (query instanceof SheetIntrospectionQuery.GetHyperlinks) {
      return new SheetInspectionResult.HyperlinksResult(
          stepId, new SheetLabel(sheetName), List.of());
    }
    return new SheetInspectionResult.CommentsResult(stepId, new SheetLabel(sheetName), List.of());
  }

  record ResolvedInspectionTarget(
      @Nullable Selector selector, @Nullable InspectionResult shortCircuitResult) {
    ResolvedInspectionTarget {
      if (selector == null && shortCircuitResult == null) {
        throw new IllegalArgumentException(
            "resolved inspection target requires either a selector or a short-circuit result");
      }
    }

    static ResolvedInspectionTarget direct(Selector selector) {
      return new ResolvedInspectionTarget(
          Objects.requireNonNull(selector, "selector must not be null"), null);
    }

    static ResolvedInspectionTarget shortCircuit(InspectionResult result) {
      return new ResolvedInspectionTarget(
          null, Objects.requireNonNull(result, "result must not be null"));
    }

    boolean isShortCircuit() {
      return shortCircuitResult != null;
    }
  }
}
