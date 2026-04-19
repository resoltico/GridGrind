package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
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
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Shared step-native workflow builders for Jazzer harnesses. */
final class ProtocolStepSupport {
  private ProtocolStepSupport() {}

  record PendingMutation(Selector target, MutationAction action) {
    PendingMutation {
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(action, "action must not be null");
    }
  }

  record PendingAssertion(String stepId, Selector target, Assertion assertion) {
    PendingAssertion {
      Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(assertion, "assertion must not be null");
    }
  }

  static PendingMutation mutate(Selector target, MutationAction action) {
    return new PendingMutation(target, action);
  }

  static PendingMutation mutate(MutationAction.SetTable action) {
    TableInput table = action.table();
    return mutate(new TableSelector.ByNameOnSheet(table.name(), table.sheetName()), action);
  }

  static PendingMutation mutate(MutationAction.SetPivotTable action) {
    PivotTableInput pivotTable = action.pivotTable();
    return mutate(
        new PivotTableSelector.ByNameOnSheet(pivotTable.name(), pivotTable.sheetName()), action);
  }

  static PendingMutation mutate(MutationAction.SetNamedRange action) {
    return mutate(namedRangeSelector(action.name(), action.scope()), action);
  }

  static PendingAssertion assertThat(String stepId, Selector target, Assertion assertion) {
    return new PendingAssertion(stepId, target, assertion);
  }

  static InspectionStep inspect(String stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(stepId, target, query);
  }

  static CalculationPolicyInput calculateAll() {
    return new CalculationPolicyInput(new CalculationStrategyInput.EvaluateAll());
  }

  static CalculationPolicyInput calculateAllAndMarkRecalculateOnOpen() {
    return new CalculationPolicyInput(new CalculationStrategyInput.EvaluateAll(), true);
  }

  static CalculationPolicyInput calculateTargets(CellSelector.QualifiedAddress... cells) {
    return new CalculationPolicyInput(new CalculationStrategyInput.EvaluateTargets(List.of(cells)));
  }

  static CalculationPolicyInput clearFormulaCaches() {
    return new CalculationPolicyInput(new CalculationStrategyInput.ClearCachesOnly());
  }

  static CalculationPolicyInput markRecalculateOnOpen() {
    return new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true);
  }

  static ExecutionPolicyInput executionPolicy(CalculationPolicyInput calculation) {
    return new ExecutionPolicyInput(null, null, calculation);
  }

  static MutationStep materializeMutation(PendingMutation mutation, int stepIndex) {
    Objects.requireNonNull(mutation, "mutation must not be null");
    if (stepIndex < 0) {
      throw new IllegalArgumentException("stepIndex must be non-negative");
    }
    return new MutationStep(
        stepIdFor(stepIndex, mutation.action()), mutation.target(), mutation.action());
  }

  static AssertionStep materializeAssertion(PendingAssertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return new AssertionStep(assertion.stepId(), assertion.target(), assertion.assertion());
  }

  static List<WorkbookStep> steps(
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    List<PendingMutation> mutationCopy = mutations == null ? List.of() : List.copyOf(mutations);
    List<PendingAssertion> assertionCopy = assertions == null ? List.of() : List.copyOf(assertions);
    List<InspectionStep> inspectionCopy =
        inspections == null ? List.of() : List.copyOf(inspections);
    List<WorkbookStep> steps =
        new ArrayList<>(mutationCopy.size() + assertionCopy.size() + inspectionCopy.size());
    for (int i = 0; i < mutationCopy.size(); i++) {
      steps.add(materializeMutation(mutationCopy.get(i), i));
    }
    assertionCopy.stream().map(ProtocolStepSupport::materializeAssertion).forEach(steps::add);
    steps.addAll(inspectionCopy);
    return List.copyOf(steps);
  }

  static List<WorkbookStep> steps(
      List<PendingMutation> mutations, List<InspectionStep> inspections) {
    return steps(mutations, List.of(), inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return new WorkbookPlan(source, persistence, steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return new WorkbookPlan(
        GridGrindProtocolVersion.current(),
        null,
        source,
        persistence,
        execution,
        formulaEnvironment,
        steps(mutations, assertions, inspections));
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return request(
        source, persistence, execution, formulaEnvironment, mutations, List.of(), inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return new WorkbookPlan(source, persistence, steps(mutations, inspections));
  }

  static int inspectionCount(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return request.inspectionSteps().size();
  }

  static List<String> inspectionIds(List<? extends InspectionResult> inspections) {
    Objects.requireNonNull(inspections, "inspections must not be null");
    return inspections.stream().map(InspectionResult::stepId).toList();
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
