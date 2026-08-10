package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.json.ActionableShapeMessage;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.Optional;
import tools.jackson.core.JsonParser;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ObjectNode;

/** Owns target-selector validation and deserialization for workbook steps. */
final class WorkbookStepJsonTargetSupport {
  private WorkbookStepJsonTargetSupport() {}

  static Selector deserializeTarget(JsonNode targetNode, JsonParser parser, String fieldName) {
    String typeFieldName = fieldName + ".type";
    if (!(targetNode instanceof ObjectNode targetObject)) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          fieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              parser,
              new ActionableShapeMessage(
                  "Field '%s' must be a JSON object".formatted(fieldName),
                  "Replace the target selector with a JSON object matching one allowed target"
                      + " shape.",
                  Optional.empty())));
    }
    String authoredType = requiredTargetType(targetObject, parser, typeFieldName);
    if (!SelectorJsonSupport.isKnownTypeId(authoredType)) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          typeFieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              parser,
              new ActionableShapeMessage(
                  "Unknown type value '%s'".formatted(authoredType),
                  "Replace field '%s' with one shipped target selector id."
                      .formatted(typeFieldName),
                  Optional.of(typeFieldName))));
    }
    return WorkbookStepJsonDeserializer.deserializeField(
        targetNode,
        parser,
        castSelectorType(SelectorJsonSupport.typeFor(authoredType).orElseThrow()),
        fieldName);
  }

  @SuppressWarnings("unchecked")
  static Class<Selector> castSelectorType(Class<? extends Selector> selectorType) {
    return (Class<Selector>) selectorType;
  }

  private static String requiredTargetType(
      ObjectNode targetNode, JsonParser parser, String typeFieldName) {
    JsonNode typeNode = targetNode.get("type");
    if (typeNode == null) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          typeFieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              parser, new MissingTypeDiscriminator(typeFieldName)));
    }
    if (!typeNode.isString()) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          typeFieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              parser,
              new ActionableShapeMessage(
                  "Field '%s' must be a string".formatted(typeFieldName),
                  "Replace field '%s' with a JSON string selector id.".formatted(typeFieldName),
                  Optional.of(typeFieldName))));
    }
    return typeNode.asString();
  }
}
