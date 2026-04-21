package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * Canonical subtype-id resolver for the public GridGrind wire contract.
 *
 * <p>The owning truth for sealed wire ids is the {@code @JsonSubTypes} annotation on the root
 * contract family. Downstream catalog, help, diagnostics, and validation surfaces derive ids from
 * that annotation map instead of duplicating string literals.
 */
public final class GridGrindProtocolTypeNames {
  private static final Map<Class<?>, String> WORKBOOK_SOURCE_TYPE_NAMES =
      typeNamesByClass(WorkbookPlan.WorkbookSource.class);
  private static final Map<Class<?>, String> WORKBOOK_PERSISTENCE_TYPE_NAMES =
      typeNamesByClass(WorkbookPlan.WorkbookPersistence.class);
  private static final Map<Class<?>, String> WORKBOOK_STEP_TYPE_NAMES =
      Map.of(
          MutationStep.class, "MUTATION",
          AssertionStep.class, "ASSERTION",
          InspectionStep.class, "INSPECTION");
  private static final Map<Class<?>, String> MUTATION_ACTION_TYPE_NAMES =
      typeNamesByClass(MutationAction.class);
  private static final Map<Class<?>, String> ASSERTION_TYPE_NAMES =
      typeNamesByClass(Assertion.class);
  private static final Map<Class<?>, String> INSPECTION_QUERY_TYPE_NAMES =
      typeNamesByClass(InspectionQuery.class);
  private static final Map<Class<?>, String> CALCULATION_STRATEGY_TYPE_NAMES =
      typeNamesByClass(CalculationStrategyInput.class);

  private GridGrindProtocolTypeNames() {}

  /** Returns the canonical wire id for one workbook-source subtype. */
  public static String workbookSourceTypeName(
      Class<? extends WorkbookPlan.WorkbookSource> workbookSourceClass) {
    return requiredTypeName(
        WORKBOOK_SOURCE_TYPE_NAMES, workbookSourceClass, "WorkbookPlan.WorkbookSource");
  }

  /** Returns the canonical wire id for one workbook-persistence subtype. */
  public static String workbookPersistenceTypeName(
      Class<? extends WorkbookPlan.WorkbookPersistence> workbookPersistenceClass) {
    return requiredTypeName(
        WORKBOOK_PERSISTENCE_TYPE_NAMES,
        workbookPersistenceClass,
        "WorkbookPlan.WorkbookPersistence");
  }

  /** Returns the canonical wire id for one workbook-step subtype. */
  public static String workbookStepTypeName(Class<? extends WorkbookStep> workbookStepClass) {
    return requiredTypeName(WORKBOOK_STEP_TYPE_NAMES, workbookStepClass, "WorkbookStep");
  }

  /** Returns the canonical wire id for one mutation-action subtype. */
  public static String mutationActionTypeName(Class<? extends MutationAction> mutationActionClass) {
    return requiredTypeName(MUTATION_ACTION_TYPE_NAMES, mutationActionClass, "MutationAction");
  }

  /** Returns the canonical wire id for one assertion subtype. */
  public static String assertionTypeName(Class<? extends Assertion> assertionClass) {
    return requiredTypeName(ASSERTION_TYPE_NAMES, assertionClass, "Assertion");
  }

  /** Returns the canonical wire id for one inspection-query subtype. */
  public static String inspectionQueryTypeName(
      Class<? extends InspectionQuery> inspectionQueryClass) {
    return requiredTypeName(INSPECTION_QUERY_TYPE_NAMES, inspectionQueryClass, "InspectionQuery");
  }

  /** Returns the canonical wire id for one calculation-strategy subtype. */
  public static String calculationStrategyTypeName(
      Class<? extends CalculationStrategyInput> calculationStrategyClass) {
    return requiredTypeName(
        CALCULATION_STRATEGY_TYPE_NAMES, calculationStrategyClass, "CalculationStrategyInput");
  }

  /** Returns the annotation-owned subtype id map for one sealed wire family. */
  static Map<Class<?>, String> typeNamesByClass(Class<?> rootType) {
    Objects.requireNonNull(rootType, "rootType must not be null");
    JsonSubTypes jsonSubTypes = rootType.getAnnotation(JsonSubTypes.class);
    if (jsonSubTypes == null) {
      throw new IllegalArgumentException(rootType + " is missing @JsonSubTypes");
    }
    return Arrays.stream(jsonSubTypes.value())
        .collect(Collectors.toUnmodifiableMap(JsonSubTypes.Type::value, JsonSubTypes.Type::name));
  }

  private static String requiredTypeName(
      Map<Class<?>, String> typeNamesByClass, Class<?> typeClass, String rootTypeName) {
    Objects.requireNonNull(typeClass, "typeClass must not be null");
    String typeName = typeNamesByClass.get(typeClass);
    if (typeName == null) {
      throw new IllegalArgumentException(
          "No discriminator name registered for " + rootTypeName + " subtype " + typeClass);
    }
    return typeName;
  }
}
