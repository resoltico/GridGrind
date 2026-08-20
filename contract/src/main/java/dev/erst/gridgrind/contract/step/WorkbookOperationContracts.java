package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.AnalysisAssertion;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.CompositeAssertion;
import dev.erst.gridgrind.contract.catalog.ProtocolTargetingMode;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Canonical per-operation contract registry used by static validation and catalog projection.
 *
 * <p>The record leaf remains the source of its public identity and summary. This registry turns the
 * leaf's operation facts into one executable contract object instead of allowing the binder,
 * catalog, and executor to each interpret targeting independently.
 */
public final class WorkbookOperationContracts {
  private static final Map<Class<? extends Record>, WorkbookOperationContract> CONTRACTS =
      new ConcurrentHashMap<>();

  private WorkbookOperationContracts() {}

  /** Returns accepted target selector types for one bound workbook operation. */
  public static Class<? extends Selector>[] targetSelectorsFor(Object operation) {
    Object boundOperation = Objects.requireNonNull(operation, "operation must not be null");
    return contractFor(recordType(boundOperation.getClass())).acceptedSelectors(boundOperation);
  }

  /** Returns the static target compatibility failure for one bound workbook operation. */
  public static java.util.Optional<String> targetViolation(Object operation, Selector target) {
    Object boundOperation = Objects.requireNonNull(operation, "operation must not be null");
    return contractFor(recordType(boundOperation.getClass()))
        .targetViolation(
            boundOperation,
            target,
            ProtocolTypeMetadataSupport.requiredTypeId(boundOperation.getClass()));
  }

  /** Returns the static execution-mode incompatibility for one bound operation. */
  public static java.util.Optional<String> executionModeViolation(
      Object operation, ExecutionModeInput executionMode) {
    Object boundOperation = Objects.requireNonNull(operation, "operation must not be null");
    return contractFor(recordType(boundOperation.getClass()))
        .executionModeViolation(
            Objects.requireNonNull(executionMode, "executionMode must not be null"));
  }

  /** Returns the catalog target surface for one concrete workbook operation record. */
  public static WorkbookStepTargeting.TargetSurface targetSurfaceFor(Class<?> operationType) {
    return contractFor(recordType(operationType)).targetSurface();
  }

  /** Returns static selector families for a concrete operation record. */
  public static Class<? extends Selector>[] staticTargetSelectorsFor(Class<?> operationType) {
    WorkbookOperationContract contract = contractFor(recordType(operationType));
    WorkbookStepTargeting.TargetSurface targetSurface = contract.targetSurface();
    if (targetSurface.rule().isPresent()) {
      throw new IllegalArgumentException(
          "Operation type " + operationType.getName() + " derives target selectors dynamically");
    }
    return contract.acceptedSelectors(operationType);
  }

  private static WorkbookOperationContract contractFor(Class<? extends Record> operationType) {
    return CONTRACTS.computeIfAbsent(operationType, WorkbookOperationContracts::build);
  }

  @SuppressWarnings("unchecked")
  private static Class<? extends Record> recordType(Class<?> operationType) {
    Objects.requireNonNull(operationType, "operationType must not be null");
    if (!operationType.isRecord()
        || !(MutationAction.class.isAssignableFrom(operationType)
            || InspectionQuery.class.isAssignableFrom(operationType)
            || Assertion.class.isAssignableFrom(operationType))) {
      throw new IllegalArgumentException("Operation type must be a record: " + operationType);
    }
    return (Class<? extends Record>) operationType;
  }

  private static WorkbookOperationContract build(Class<? extends Record> operationType) {
    ProtocolTargetingMode mode = ProtocolTypeMetadataSupport.targetingMode(operationType);
    WorkbookOperationTargetContract targetSelectorContract =
        switch (mode) {
          case STATIC ->
              new StaticWorkbookOperationTargetContract(
                  Arrays.asList(ProtocolTypeMetadataSupport.staticTargetSelectors(operationType)));
          case ANALYSIS_QUERY ->
              derived(
                  AnalysisAssertion.ANALYSIS_RULE,
                  operation ->
                      AnalysisAssertion.targetSelectorsFor(
                          AnalysisAssertion.class.cast(operation)));
          case NESTED_ASSERTION ->
              derived(
                  CompositeAssertion.NESTED_RULE,
                  operation ->
                      CompositeAssertion.targetSelectorsFor(
                          CompositeAssertion.class.cast(operation)));
          case INTERSECTION_OF_NESTED_ASSERTIONS ->
              derived(
                  CompositeAssertion.INTERSECTION_RULE,
                  operation ->
                      CompositeAssertion.targetSelectorsFor(
                          CompositeAssertion.class.cast(operation)));
        };
    return new WorkbookOperationContract(operationType, targetSelectorContract);
  }

  private static WorkbookOperationTargetContract derived(
      String rule, WorkbookOperationTargetSelectorDerivation derivation) {
    return new DerivedWorkbookOperationTargetContract(rule, derivation);
  }
}
