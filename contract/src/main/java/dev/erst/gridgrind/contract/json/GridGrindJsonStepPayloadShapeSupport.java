package dev.erst.gridgrind.contract.json;

import java.util.Optional;
import tools.jackson.databind.JsonNode;

/** Owns raw-tree validation for the mutually exclusive step payload families. */
final class GridGrindJsonStepPayloadShapeSupport {
  private GridGrindJsonStepPayloadShapeSupport() {}

  static void rejectInvalidStepPayloadShapes(JsonNode node) {
    if (!node.isObject()) {
      return;
    }
    JsonNode stepsNode = node.get("steps");
    if (stepsNode == null || !stepsNode.isArray()) {
      return;
    }
    for (int index = 0; index < stepsNode.size(); index++) {
      rejectInvalidStepPayloadShape(stepsNode.get(index), index);
    }
  }

  static String conflictingPayloadField(JsonNode stepNode) {
    boolean actionPresent = stepNode.has("action");
    boolean assertionPresent = stepNode.has("assertion");
    boolean queryPresent = stepNode.has("query");
    if (actionPresent && assertionPresent) {
      return "assertion";
    }
    if (actionPresent && queryPresent) {
      return "query";
    }
    if (assertionPresent && queryPresent) {
      return "query";
    }
    throw new IllegalStateException("conflictingPayloadField requires at least two payloads");
  }

  private static void rejectInvalidStepPayloadShape(JsonNode stepNode, int index) {
    if (!stepNode.isObject()) {
      return;
    }
    int payloadCount = payloadCount(stepNode);
    if (payloadCount == 0) {
      throw invalidStepPayloadShape("steps[" + index + "]");
    }
    if (payloadCount > 1) {
      throw invalidStepPayloadShape("steps[" + index + "]." + conflictingPayloadField(stepNode));
    }
  }

  private static int payloadCount(JsonNode stepNode) {
    return (stepNode.has("action") ? 1 : 0)
        + (stepNode.has("assertion") ? 1 : 0)
        + (stepNode.has("query") ? 1 : 0);
  }

  private static InvalidRequestShapeException invalidStepPayloadShape(String jsonPath) {
    return new InvalidRequestShapeException(
        new MessageShape(
            "Each step must contain exactly one of 'action', 'assertion', or 'query'",
            Optional.of(jsonPath)),
        Optional.of(jsonPath),
        Optional.empty(),
        Optional.empty(),
        null);
  }
}
