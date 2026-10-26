package dev.erst.gridgrind.engine.runtime.parity;

import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
    TableSpec table = action.table();
    return mutate(new TableSelector.ByNameOnSheet(table.name(), table.sheetName()), action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetPivotTable action) {
    PivotTableSpec pivotTable = action.pivotTable();
    return mutate(
        new PivotTableSelector.ByNameOnSheet(pivotTable.name(), pivotTable.sheetName()), action);
  }

  static PendingMutation mutate(StructuredMutationAction.SetNamedRange action) {
    return mutate(namedRangeSelector(action.name().value(), action.scope()), action);
  }

  static InspectionStep inspect(String stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(new StepId(stepId), target, query);
  }

  static CalculationPolicySpec calculateAll() {
    return CalculationPolicySpec.strategy(new CalculationStrategySpec.EvaluateAll());
  }

  static CalculationPolicySpec calculateAllAndMarkRecalculateOnOpen() {
    return new CalculationPolicySpec(new CalculationStrategySpec.EvaluateAll(), true);
  }

  static CalculationPolicySpec calculateTargets(
      dev.erst.gridgrind.core.model.formula.FormulaCellTarget... cells) {
    return CalculationPolicySpec.strategy(
        new CalculationStrategySpec.EvaluateTargets(List.of(cells)));
  }

  static CalculationPolicySpec clearFormulaCaches() {
    return CalculationPolicySpec.strategy(new CalculationStrategySpec.ClearCachesOnly());
  }

  static CalculationPolicySpec markRecalculateOnOpen() {
    return new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), true);
  }

  static ExecutionPolicySpec executionPolicy(CalculationPolicySpec calculation) {
    return ExecutionPolicySpec.calculation(calculation);
  }

  static ExecutionPolicySpec executionPolicy(
      dev.erst.gridgrind.core.plan.ExecutionModeSpec mode, CalculationPolicySpec calculation) {
    return ExecutionPolicySpec.modeAndCalculation(mode, calculation);
  }

  static MutationStep materializeMutation(PendingMutation mutation, int stepIndex) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    if (stepIndex < 0) {
      throw new IllegalArgumentException("stepIndex must be non-negative");
    }
    return new MutationStep(
        new StepId(stepIdFor(stepIndex, mutation.action())), mutation.target(), mutation.action());
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
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicySpec execution,
      FormulaEnvironmentSpec formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return WorkbookPlan.standard(
        source,
        persistence,
        Objects.requireNonNullElseGet(execution, ExecutionPolicySpec::defaults),
        Objects.requireNonNullElseGet(formulaEnvironment, FormulaEnvironmentSpec::empty),
        steps(mutations, inspections));
  }

  static <T extends InspectionResult> T inspection(
      List<InspectionResult> inspections, String stepId, Class<T> type) {
    Objects.requireNonNull(inspections, "inspections must not be null");
    Objects.requireNonNull(stepId, "stepId must not be null");
    Objects.requireNonNull(type, "type must not be null");
    return type.cast(
        inspections.stream()
            .filter(result -> result.stepId().value().equals(stepId))
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
      case NamedRangeScope.Workbook _ ->
          new NamedRangeSelector.WorkbookScope(new DefinedName(name));
      case NamedRangeScope.Sheet sheet ->
          new NamedRangeSelector.SheetScope(new DefinedName(name), sheet.sheetName());
    };
  }
}
