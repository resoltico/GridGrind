package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.json.ActionableShapeMessage;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import tools.jackson.core.JsonParser;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ObjectNode;

/** Owns target-selector validation and deserialization for workbook steps. */
final class WorkbookStepJsonTargetSupport {
  private WorkbookStepJsonTargetSupport() {}

  @SafeVarargs
  static Selector deserializeTarget(
      JsonNode targetNode,
      JsonParser parser,
      String fieldName,
      Class<? extends Selector>... allowedTypes) {
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
    String allowedSummary = selectorFamilySummary(Arrays.asList(allowedTypes));
    if (!SelectorJsonSupport.isKnownTypeId(authoredType)) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          typeFieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              parser,
              new ActionableShapeMessage(
                  unknownTargetTypeMessage(typeFieldName, authoredType, allowedSummary),
                  "Replace field '%s' with one of the allowed target selector ids for this step."
                      .formatted(typeFieldName),
                  Optional.of(typeFieldName))));
    }
    Set<String> allowedTypeIds = new LinkedHashSet<>();
    Class<? extends Selector> candidateType = null;
    for (Class<? extends Selector> allowedType : allowedTypes) {
      List<String> selectorTypeIds = SelectorJsonSupport.typeIdsFor(allowedType);
      allowedTypeIds.addAll(selectorTypeIds);
      if (candidateType == null && selectorTypeIds.contains(authoredType)) {
        candidateType = allowedType;
      }
    }
    if (!allowedTypeIds.contains(authoredType)) {
      throw WorkbookStepJsonFailurePathSupport.fieldFailure(
          typeFieldName,
          WorkbookStepJsonFailurePathSupport.inputMismatch(
              parser,
              new ActionableShapeMessage(
                  ("Field '%s' uses target selector type '%s', which is not allowed for this step;"
                          + " allowed targets: %s")
                      .formatted(typeFieldName, authoredType, allowedSummary),
                  "Replace field '%s' with one of the allowed target selector ids for this step."
                      .formatted(typeFieldName),
                  Optional.of(typeFieldName))));
    }
    return WorkbookStepJsonDeserializer.deserializeField(
        targetNode,
        parser,
        castSelectorType(
            Objects.requireNonNull(
                candidateType,
                "Selector type ids must be globally unique per step target; authored type '%s'"
                    .formatted(authoredType))),
        fieldName);
  }

  @SuppressWarnings("unchecked")
  static Class<Selector> castSelectorType(Class<? extends Selector> selectorType) {
    return (Class<Selector>) selectorType;
  }

  static String selectorFamilySummary(Iterable<Class<? extends Selector>> selectorTypes) {
    return SelectorJsonSupport.familySummary(selectorTypes);
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

  private static String unknownTargetTypeMessage(
      String typeFieldName, String authoredType, String allowedTargets) {
    String guidance = WorkbookStepLegacySelectorTypeHints.guidancePrefix(authoredType);
    return "Field '%s' uses unknown target selector type '%s'; %sallowed targets: %s"
        .formatted(typeFieldName, authoredType, guidance, allowedTargets);
  }
}
