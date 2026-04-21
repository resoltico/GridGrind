package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JacksonException.Reference;
import tools.jackson.core.JsonParser;
import tools.jackson.databind.DatabindException;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.ValueDeserializer;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.node.ObjectNode;

/** Deserializes workbook steps from the canonical step envelope without a redundant outer type. */
final class WorkbookStepJsonDeserializer extends ValueDeserializer<WorkbookStep> {
  private static final Set<String> ALLOWED_FIELDS =
      Set.of("stepId", "target", "action", "assertion", "query");

  @Override
  public WorkbookStep deserialize(JsonParser parser, DeserializationContext context) {
    JsonNode rawNode = parser.readValueAsTree();
    if (!(rawNode instanceof ObjectNode stepNode)) {
      throw inputMismatch(parser, "steps entries must be JSON objects");
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
      throw inputMismatch(
          parser, "Each step must contain exactly one of 'action', 'assertion', or 'query'");
    }
    if (actionNode != null) {
      MutationAction action =
          deserializeField(actionNode, parser, context, MutationAction.class, "action");
      Selector target =
          deserializeTarget(
              targetNode,
              parser,
              context,
              "target",
              WorkbookStepValidation.allowedTargetTypes(action));
      return new MutationStep(stepId, target, action);
    }
    if (assertionNode != null) {
      Assertion assertion =
          deserializeField(assertionNode, parser, context, Assertion.class, "assertion");
      Selector target =
          deserializeTarget(
              targetNode,
              parser,
              context,
              "target",
              WorkbookStepValidation.allowedTargetTypes(assertion));
      return new AssertionStep(stepId, target, assertion);
    }
    InspectionQuery query =
        deserializeField(queryNode, parser, context, InspectionQuery.class, "query");
    Selector target =
        deserializeTarget(
            targetNode,
            parser,
            context,
            "target",
            WorkbookStepValidation.allowedTargetTypes(query));
    return new InspectionStep(stepId, target, query);
  }

  private static void rejectUnknownFields(ObjectNode stepNode, DeserializationContext context) {
    Iterator<String> fields = stepNode.propertyNames().iterator();
    while (fields.hasNext()) {
      String fieldName = fields.next();
      if (!ALLOWED_FIELDS.contains(fieldName)) {
        throw inputMismatch(context.getParser(), "Unknown field '%s'".formatted(fieldName));
      }
    }
  }

  private static JsonNode requiredNode(
      ObjectNode stepNode, String fieldName, DeserializationContext context) {
    JsonNode node = stepNode.get(fieldName);
    if (node == null) {
      throw inputMismatch(context.getParser(), "Missing required field '%s'".formatted(fieldName));
    }
    return node;
  }

  private static String requiredText(
      ObjectNode stepNode, String fieldName, DeserializationContext context) {
    JsonNode node = requiredNode(stepNode, fieldName, context);
    if (!node.isString()) {
      throw inputMismatch(context.getParser(), "Field '%s' must be a string".formatted(fieldName));
    }
    return node.asString();
  }

  private static <T> T deserializeNode(JsonNode node, JsonParser parser, Class<T> targetType) {
    try (JsonParser nodeParser = node.traverse(parser.objectReadContext())) {
      return nodeParser.readValueAs(targetType);
    }
  }

  private static MismatchedInputException inputMismatch(JsonParser parser, String message) {
    return MismatchedInputException.from(parser, WorkbookStep.class, message);
  }

  private static <T> T deserializeField(
      JsonNode node,
      JsonParser parser,
      DeserializationContext context,
      Class<T> targetType,
      String fieldName) {
    try {
      return deserializeNode(node, parser, targetType);
    } catch (JacksonException exception) {
      throw DatabindException.wrapWithPath(
          context, exception, new Reference(WorkbookStep.class, fieldName));
    }
  }

  @SafeVarargs
  private static Selector deserializeTarget(
      JsonNode targetNode,
      JsonParser parser,
      DeserializationContext context,
      String fieldName,
      Class<? extends Selector>... allowedTypes) {
    if (!(targetNode instanceof ObjectNode targetObject)) {
      throw fieldFailure(
          context,
          fieldName,
          inputMismatch(parser, "Field '%s' must be a JSON object".formatted(fieldName)));
    }
    String authoredType = requiredTargetType(targetObject, parser, context, fieldName);
    if (!SelectorJsonSupport.isKnownTypeId(authoredType)) {
      throw fieldFailure(
          context,
          fieldName,
          inputMismatch(parser, "Unknown type value '%s'".formatted(authoredType)));
    }
    Set<String> allowedTypeIds =
        Arrays.stream(allowedTypes)
            .flatMap(type -> SelectorJsonSupport.typeIdsFor(type).stream())
            .collect(java.util.stream.Collectors.toCollection(LinkedHashSet::new));
    if (!allowedTypeIds.contains(authoredType)) {
      throw fieldFailure(
          context,
          fieldName,
          inputMismatch(
              parser,
              "Target selector type '%s' is not allowed for this step; allowed targets: %s"
                  .formatted(authoredType, selectorFamilySummary(Arrays.asList(allowedTypes)))));
    }
    List<Class<? extends Selector>> candidateTypes =
        Arrays.stream(allowedTypes)
            .filter(allowedType -> SelectorJsonSupport.supportsTypeId(allowedType, authoredType))
            .toList();
    if (candidateTypes.size() == 1) {
      return deserializeField(
          targetNode, parser, context, castSelectorType(candidateTypes.getFirst()), fieldName);
    }
    return deserializeAmbiguousTarget(
        targetNode, parser, context, fieldName, authoredType, candidateTypes);
  }

  private static String requiredTargetType(
      ObjectNode targetNode, JsonParser parser, DeserializationContext context, String fieldName) {
    JsonNode typeNode = targetNode.get("type");
    if (typeNode == null) {
      throw fieldFailure(
          context, fieldName, inputMismatch(parser, "Missing required field 'type'"));
    }
    if (!typeNode.isString()) {
      throw fieldFailure(
          context, fieldName, inputMismatch(parser, "Field 'type' must be a string"));
    }
    return typeNode.asString();
  }

  @SuppressWarnings("unchecked")
  static Class<Selector> castSelectorType(Class<? extends Selector> selectorType) {
    return (Class<Selector>) selectorType;
  }

  private static Selector deserializeAmbiguousTarget(
      JsonNode targetNode,
      JsonParser parser,
      DeserializationContext context,
      String fieldName,
      String authoredType,
      List<Class<? extends Selector>> candidateTypes) {
    Selector resolvedSelector = null;
    List<Class<? extends Selector>> successfulTypes = new java.util.ArrayList<>();
    for (Class<? extends Selector> candidateType : candidateTypes) {
      try {
        Selector candidate = deserializeNode(targetNode, parser, castSelectorType(candidateType));
        if (resolvedSelector == null) {
          resolvedSelector = candidate;
        }
        successfulTypes.add(candidateType);
      } catch (JacksonException ignored) {
        // Deliberately try the other selector families that share the same wire type.
      }
    }
    if (successfulTypes.size() == 1) {
      return resolvedSelector;
    }
    if (successfulTypes.size() > 1) {
      throw fieldFailure(
          context,
          fieldName,
          inputMismatch(
              parser,
              ("Target selector type '%s' is ambiguous for this step; the authored object"
                      + " matches multiple selector families: %s")
                  .formatted(
                      authoredType, matchingSelectorFamilySummary(authoredType, successfulTypes))));
    }
    throw fieldFailure(
        context,
        fieldName,
        inputMismatch(
            parser,
            ("Target selector type '%s' has the wrong shape for this step; none of the matching"
                    + " selector families accepted the authored fields. Matching families: %s")
                .formatted(
                    authoredType, matchingSelectorFamilySummary(authoredType, candidateTypes))));
  }

  private static JacksonException fieldFailure(
      DeserializationContext context, String fieldName, JacksonException failure) {
    return DatabindException.wrapWithPath(
        context, failure, new Reference(WorkbookStep.class, fieldName));
  }

  static String selectorFamilySummary(Iterable<Class<? extends Selector>> selectorTypes) {
    return SelectorJsonSupport.familySummary(selectorTypes);
  }

  private static String matchingSelectorFamilySummary(
      String authoredType, Iterable<Class<? extends Selector>> selectorTypes) {
    Set<String> familyNames = new LinkedHashSet<>();
    for (Class<? extends Selector> selectorType : selectorTypes) {
      familyNames.add(SelectorJsonSupport.familyName(selectorType));
    }
    return familyNames.stream()
        .map(familyName -> familyName + "(" + authoredType + ")")
        .collect(java.util.stream.Collectors.joining("; "));
  }
}
