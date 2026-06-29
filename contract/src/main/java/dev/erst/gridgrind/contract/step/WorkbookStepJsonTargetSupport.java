package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonParser;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.node.ObjectNode;

/** Owns target-selector validation and deserialization for workbook steps. */
final class WorkbookStepJsonTargetSupport {
  // These ids are used only to tailor the error message when a caller uses retired generic names.
  private static final Set<String> RENAMED_SELECTOR_TYPE_HINTS =
      Set.of(
          "CURRENT",
          "ALL",
          "ALL_ON_SHEET",
          "ALL_ROWS",
          "ALL_USED_IN_SHEET",
          "ANY_OF",
          "BY_ADDRESS",
          "BY_ADDRESSES",
          "BY_COLUMN_NAME",
          "BY_INDEX",
          "BY_KEY_CELL",
          "BY_NAME",
          "BY_NAME_ON_SHEET",
          "BY_NAMES",
          "BY_QUALIFIED_ADDRESSES",
          "BY_RANGE",
          "BY_RANGES",
          "INSERTION",
          "RECTANGULAR_WINDOW",
          "SHEET_SCOPE",
          "SPAN",
          "WORKBOOK_SCOPE");

  private WorkbookStepJsonTargetSupport() {}

  @SafeVarargs
  static Selector deserializeTarget(
      JsonNode targetNode,
      JsonParser parser,
      String fieldName,
      Class<? extends Selector>... allowedTypes) {
    if (!(targetNode instanceof ObjectNode targetObject)) {
      throw fieldFailure(
          fieldName,
          inputMismatch(parser, "Field '%s' must be a JSON object".formatted(fieldName)));
    }
    String authoredType = requiredTargetType(targetObject, parser, fieldName);
    if (!SelectorJsonSupport.isKnownTypeId(authoredType)) {
      throw fieldFailure(
          fieldName + ".type",
          inputMismatch(
              parser,
              unknownTargetTypeMessage(
                  authoredType, selectorFamilySummary(Arrays.asList(allowedTypes)))));
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
      throw fieldFailure(
          fieldName + ".type",
          inputMismatch(
              parser,
              "Target selector type '%s' is not allowed for this step; allowed targets: %s"
                  .formatted(authoredType, selectorFamilySummary(Arrays.asList(allowedTypes)))));
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
      ObjectNode targetNode, JsonParser parser, String fieldName) {
    JsonNode typeNode = targetNode.get("type");
    if (typeNode == null) {
      throw fieldFailure(
          fieldName + ".type", inputMismatch(parser, "Missing required field 'type'"));
    }
    if (!typeNode.isString()) {
      throw fieldFailure(
          fieldName + ".type", inputMismatch(parser, "Field 'type' must be a string"));
    }
    return typeNode.asString();
  }

  private static String unknownTargetTypeMessage(String authoredType, String allowedTargets) {
    String guidance =
        RENAMED_SELECTOR_TYPE_HINTS.contains(authoredType)
            ? "target selector ids are family-specific; "
            : "";
    return "Unknown target selector type '%s'; %sallowed targets: %s"
        .formatted(authoredType, guidance, allowedTargets);
  }

  private static MismatchedInputException inputMismatch(JsonParser parser, String message) {
    return MismatchedInputException.from(parser, WorkbookStep.class, message);
  }

  private static JacksonException fieldFailure(String fieldName, JacksonException failure) {
    return failure.prependPath(WorkbookStep.class, fieldName);
  }
}
