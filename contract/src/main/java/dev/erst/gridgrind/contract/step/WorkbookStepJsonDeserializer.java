package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.json.ActionableShapeMessage;
import dev.erst.gridgrind.contract.json.MissingRequiredField;
import dev.erst.gridgrind.contract.json.UnknownField;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import java.util.Optional;
import java.util.Set;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonParser;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ValueDeserializer;
import tools.jackson.databind.node.ObjectNode;

/** Deserializes workbook steps from the canonical step envelope without a redundant outer type. */
final class WorkbookStepJsonDeserializer extends ValueDeserializer<WorkbookStep> {
  private static final Set<String> ALLOWED_FIELDS =
      Set.of("stepId", "target", "action", "assertion", "query");

  @Override
  public WorkbookStep deserialize(JsonParser parser, DeserializationContext context) {
    JsonNode rawNode = parser.readValueAsTree();
    if (!(rawNode instanceof ObjectNode stepNode)) {
      throw WorkbookStepJsonFailurePathSupport.inputMismatch(
          parser,
          new ActionableShapeMessage(
              "steps entries must be JSON objects",
              "Replace each steps entry with one JSON object containing stepId, target, and"
                  + " exactly one of action, assertion, or query.",
              Optional.empty()));
    }
    rejectUnknownFields(stepNode, context);

    String stepId = requiredText(stepNode, "stepId", context);
    JsonNode targetNode = requiredNode(stepNode, "target", context);
    JsonNode actionNode = stepNode.get("action");
    JsonNode assertionNode = stepNode.get("assertion");
    JsonNode queryNode = stepNode.get("query");
    int stepPayloadCount =
        (actionNode == null ? 0 : 1)
            + (assertionNode == null ? 0 : 1)
            + (queryNode == null ? 0 : 1);
    if (stepPayloadCount != 1) {
      throw stepFailure(
          context,
          new ActionableShapeMessage(
              "Each step must contain exactly one of 'action', 'assertion', or 'query'",
              "Add exactly one of action, assertion, or query to each step.",
              Optional.empty()));
    }
    if (actionNode != null) {
      MutationAction action = deserializeField(actionNode, parser, MutationAction.class, "action");
      Selector target =
          WorkbookStepJsonTargetSupport.deserializeTarget(
              targetNode, parser, "target", WorkbookStepValidation.allowedTargetTypes(action));
      return new MutationStep(stepId, target, action);
    }
    if (assertionNode != null) {
      Assertion assertion = deserializeField(assertionNode, parser, Assertion.class, "assertion");
      Selector target =
          WorkbookStepJsonTargetSupport.deserializeTarget(
              targetNode, parser, "target", WorkbookStepValidation.allowedTargetTypes(assertion));
      return new AssertionStep(stepId, target, assertion);
    }
    InspectionQuery query = deserializeField(queryNode, parser, InspectionQuery.class, "query");
    Selector target =
        WorkbookStepJsonTargetSupport.deserializeTarget(
            targetNode, parser, "target", WorkbookStepValidation.allowedTargetTypes(query));
    return new InspectionStep(stepId, target, query);
  }

  private static void rejectUnknownFields(ObjectNode stepNode, DeserializationContext context) {
    String fieldName = firstUnexpectedFieldName(stepNode);
    if (fieldName != null) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          fieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              context.getParser(), new UnknownField(fieldName)));
    }
  }

  private static JsonNode requiredNode(
      ObjectNode stepNode, String fieldName, DeserializationContext context) {
    JsonNode node = stepNode.get(fieldName);
    if (node == null) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          fieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              context.getParser(), new MissingRequiredField(fieldName)));
    }
    return node;
  }

  private static String requiredText(
      ObjectNode stepNode, String fieldName, DeserializationContext context) {
    JsonNode node = requiredNode(stepNode, fieldName, context);
    if (!node.isString()) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          fieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              context.getParser(),
              new ActionableShapeMessage(
                  "Field '%s' must be a string".formatted(fieldName),
                  "Replace field '%s' with a JSON string value.".formatted(fieldName),
                  Optional.empty())));
    }
    return node.asString();
  }

  private static <T> T deserializeNode(JsonNode node, JsonParser parser, Class<T> targetType) {
    try (JsonParser nodeParser = node.traverse(parser.objectReadContext())) {
      return nodeParser.readValueAs(targetType);
    }
  }

  static <T> T deserializeField(
      JsonNode node, JsonParser parser, Class<T> targetType, String fieldName) {
    try {
      return deserializeNode(node, parser, targetType);
    } catch (JacksonException exception) {
      throw WorkbookStepJsonFailurePathSupport.wrapJacksonFailure(
          fieldName, node, targetType, exception);
    } catch (IllegalArgumentException exception) {
      throw WorkbookStepJsonFailurePathSupport.wrapIllegalArgumentFailure(
          parser, fieldName, exception);
    }
  }

  private static @Nullable String firstUnexpectedFieldName(ObjectNode stepNode) {
    var fields = stepNode.propertyNames().iterator();
    while (fields.hasNext()) {
      String fieldName = fields.next();
      if (!ALLOWED_FIELDS.contains(fieldName)) {
        return fieldName;
      }
    }
    return null;
  }

  private static JacksonException stepFailure(
      DeserializationContext context, ActionableShapeMessage requestProblem) {
    return WorkbookStepJsonFailurePathSupport.inputMismatch(context.getParser(), requestProblem);
  }
}
