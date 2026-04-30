package dev.erst.gridgrind.executor.parity;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Shared step-native request builders for parity tests. */
final class ParityPlanSupport {
  private ParityPlanSupport() {}

  record PendingMutation(Selector target, MutationAction action) {
    PendingMutation {
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(action, "action must not be null");
    }
  }

  static PendingMutation mutate(Selector target, MutationAction action) {
    return new PendingMutation(target, action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetTable action) {
    TableInput table = action.table();
    return mutate(new TableSelector.ByNameOnSheet(table.name(), table.sheetName()), action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetPivotTable action) {
    PivotTableInput pivotTable = action.pivotTable();
    return mutate(
        new PivotTableSelector.ByNameOnSheet(pivotTable.name(), pivotTable.sheetName()), action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetNamedRange action) {
    return mutate(namedRangeSelector(action.name(), action.scope()), action);
  }

  static InspectionStep inspect(String stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(stepId, target, query);
  }

  static CalculationPolicyInput calculateAll() {
    return CalculationPolicyInput.strategy(new CalculationStrategyInput.EvaluateAll());
  }

  static CalculationPolicyInput calculateAllAndMarkRecalculateOnOpen() {
    return new CalculationPolicyInput(new CalculationStrategyInput.EvaluateAll(), true);
  }

  static CalculationPolicyInput calculateTargets(CellSelector.QualifiedAddress... cells) {
    return CalculationPolicyInput.strategy(
        new CalculationStrategyInput.EvaluateTargets(List.of(cells)));
  }

  static CalculationPolicyInput clearFormulaCaches() {
    return CalculationPolicyInput.strategy(new CalculationStrategyInput.ClearCachesOnly());
  }

  static CalculationPolicyInput markRecalculateOnOpen() {
    return new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true);
  }

  static ExecutionPolicyInput executionPolicy(CalculationPolicyInput calculation) {
    return ExecutionPolicyInput.calculation(calculation);
  }

  static ExecutionPolicyInput executionPolicy(
      dev.erst.gridgrind.contract.dto.ExecutionModeInput mode, CalculationPolicyInput calculation) {
    return ExecutionPolicyInput.modeAndCalculation(mode, calculation);
  }

  static MutationStep materializeMutation(PendingMutation mutation, int stepIndex) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    if (stepIndex < 0) {
      throw new IllegalArgumentException("stepIndex must be non-negative");
    }
    return new MutationStep(
        stepIdFor(stepIndex, mutation.action()), mutation.target(), mutation.action());
  }

  static List<WorkbookStep> steps(
      List<PendingMutation> mutations, List<InspectionStep> inspections) {
    List<PendingMutation> mutationCopy = mutations == null ? List.of() : List.copyOf(mutations);
    List<InspectionStep> inspectionCopy =
        inspections == null ? List.of() : List.copyOf(inspections);
    List<WorkbookStep> steps = new ArrayList<>(mutationCopy.size() + inspectionCopy.size());
    for (int i = 0; i < mutationCopy.size(); i++) {
      steps.add(materializeMutation(mutationCopy.get(i), i));
    }
    steps.addAll(inspectionCopy);
    return List.copyOf(steps);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicyInput.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        Objects.requireNonNullElseGet(execution, ExecutionPolicyInput::defaults),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentInput::empty),
        steps(mutations, inspections));
  }

  static <T extends InspectionResult> T inspection(
      List<InspectionResult> inspections, String stepId, Class<T> type) {
    Objects.requireNonNull(inspections, "inspections must not be null");
    Objects.requireNonNull(stepId, "stepId must not be null");
    Objects.requireNonNull(type, "type must not be null");
    return type.cast(
        inspections.stream()
            .filter(result -> result.stepId().equals(stepId))
            .findFirst()
            .orElseThrow(
                () -> new AssertionError("Missing inspection result for stepId " + stepId)));
  }

  private static String stepIdFor(int stepIndex, MutationAction action) {
    String normalizedType = action.actionType().toLowerCase(Locale.ROOT).replace('_', '-');
    return "step-" + String.format(Locale.ROOT, "%02d", stepIndex + 1) + "-" + normalizedType;
  }

  private static Selector namedRangeSelector(String name, NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> new NamedRangeSelector.WorkbookScope(name);
      case NamedRangeScope.Sheet sheet ->
          new NamedRangeSelector.SheetScope(name, sheet.sheetName());
    };
  }
}
