package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ObjectNode;

/** Builds safe placeholder steps so doctor mode can keep decoding later malformed steps. */
final class CliDoctorRequestPlaceholderStepFactory {
  private static final String PLACEHOLDER_SHEET_NAME = "GridGrindDoctorStep";

  private CliDoctorRequestPlaceholderStepFactory() {}

  static ObjectNode placeholderStepNode(
      JsonNode authoredStep, int stepIndex, String executionModeType) {
    WorkbookStep placeholder =
        placeholderStep(
            stepIdForPlaceholder(authoredStep, stepIndex), authoredStep, executionModeType);
    ObjectNode request =
        GridGrindJson.requestTree(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                ExecutionPolicyInput.defaults(),
                FormulaEnvironmentInput.empty(),
                List.of(placeholder)));
    return (ObjectNode) request.withArray("steps").get(0).deepCopy();
  }

  static String authoredExecutionModeType(ObjectNode root) {
    Objects.requireNonNull(root, "root must not be null");
    String authoredType = authoredScalarText(root.path("execution").path("mode").path("type"));
    return authoredType.isBlank() ? "FULL_XSSF" : authoredType;
  }

  private static WorkbookStep placeholderStep(
      String stepId, JsonNode authoredStep, String executionModeType) {
    return switch (preferredPlaceholderFamily(authoredStep, executionModeType)) {
      case MUTATION ->
          new MutationStep(
              stepId,
              new SheetSelector.ByName(PLACEHOLDER_SHEET_NAME),
              new WorkbookMutationAction.EnsureSheet());
      case ASSERTION ->
          new AssertionStep(
              stepId,
              new SheetSelector.ByName(PLACEHOLDER_SHEET_NAME),
              new PresenceAssertion.SheetPresent());
      case QUERY ->
          new InspectionStep(
              stepId,
              new WorkbookSelector.Current(),
              new WorkbookIntrospectionQuery.GetWorkbookSummary());
    };
  }

  private static String stepIdForPlaceholder(JsonNode authoredStep, int stepIndex) {
    if (!(authoredStep instanceof ObjectNode stepObject)) {
      return "gridgrind-doctor-step-" + stepIndex;
    }
    JsonNode stepIdNode = stepObject.get("stepId");
    if (stepIdNode == null || !stepIdNode.isString()) {
      return "gridgrind-doctor-step-" + stepIndex;
    }
    String candidate = stepIdNode.stringValue();
    return validPlaceholderStepId(candidate) ? candidate : "gridgrind-doctor-step-" + stepIndex;
  }

  private static boolean validPlaceholderStepId(String candidate) {
    try {
      new InspectionStep(
          candidate,
          new WorkbookSelector.Current(),
          new WorkbookIntrospectionQuery.GetWorkbookSummary());
      return true;
    } catch (IllegalArgumentException exception) {
      return false;
    }
  }

  private static PlaceholderStepFamily preferredPlaceholderFamily(
      JsonNode authoredStep, String executionModeType) {
    List<PlaceholderStepFamily> presentFamilies = new ArrayList<>(3);
    if (authoredStep instanceof ObjectNode stepObject) {
      if (stepObject.has("action")) {
        presentFamilies.add(PlaceholderStepFamily.MUTATION);
      }
      if (stepObject.has("assertion")) {
        presentFamilies.add(PlaceholderStepFamily.ASSERTION);
      }
      if (stepObject.has("query")) {
        presentFamilies.add(PlaceholderStepFamily.QUERY);
      }
    }
    if (presentFamilies.isEmpty()) {
      return defaultPlaceholderFamily(executionModeType);
    }
    return presentFamilies.stream()
        .filter(family -> familyCompatibleWithMode(family, executionModeType))
        .findFirst()
        .orElse(presentFamilies.getFirst());
  }

  private static PlaceholderStepFamily defaultPlaceholderFamily(String executionModeType) {
    if ("STREAMING_WRITE".equals(executionModeType)) {
      return PlaceholderStepFamily.MUTATION;
    }
    return PlaceholderStepFamily.QUERY;
  }

  private static boolean familyCompatibleWithMode(
      PlaceholderStepFamily family, String executionModeType) {
    if ("EVENT_READ".equals(executionModeType)) {
      return family == PlaceholderStepFamily.QUERY;
    }
    return !"STREAMING_WRITE".equals(executionModeType) || family == PlaceholderStepFamily.MUTATION;
  }

  private static String authoredScalarText(JsonNode node) {
    Objects.requireNonNull(node, "node must not be null");
    if (node.isNull() || node.isMissingNode()) {
      return "";
    }
    if (node.isString()) {
      return node.stringValue();
    }
    if (node.isNumber()) {
      return node.numberValue().toString();
    }
    if (node.isBoolean()) {
      return Boolean.toString(node.booleanValue());
    }
    return "";
  }

  /** Safe synthetic step families that preserve enough shape for continued doctor decoding. */
  private enum PlaceholderStepFamily {
    /** Synthetic mutation placeholder. */
    MUTATION,

    /** Synthetic assertion placeholder. */
    ASSERTION,

    /** Synthetic inspection placeholder. */
    QUERY
  }
}
