package dev.erst.gridgrind.contract.json;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.List;
import java.util.Optional;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.JsonNodeFactory;
import tools.jackson.databind.node.ObjectNode;

/** Converts structurally valid tolerant fragments into their independently bound DTO values. */
final class RequestFragmentBinder {
  private RequestFragmentBinder() {}

  static <T> Optional<T> bindMember(
      Optional<RequestJsonMember> member,
      Class<T> type,
      String jsonPath,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures,
      RequestDiagnosticRedactor diagnosticRedactor) {
    return member.flatMap(
        value ->
            bindNode(
                value.value(),
                type,
                jsonPath,
                structuralProblems,
                bindingFailures,
                diagnosticRedactor));
  }

  static <T> Optional<T> bindNode(
      RequestJsonNode node,
      Class<T> type,
      String jsonPath,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures,
      RequestDiagnosticRedactor diagnosticRedactor) {
    if (hasProblemAtOrBelow(jsonPath, structuralProblems) || !canConvert(node)) {
      return Optional.empty();
    }
    try {
      return Optional.of(
          GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER.treeToValue(toJsonNode(node), type));
    } catch (RuntimeException exception) {
      bindingFailures.add(
          RequestBindingFailure.from(exception, node, jsonPath, diagnosticRedactor));
      return Optional.empty();
    }
  }

  static boolean hasProblemAtOrBelow(String jsonPath, List<RequestStructuralProblem> problems) {
    return problems.stream()
        .anyMatch(
            problem -> {
              if (problem instanceof RequestDuplicateKey duplicateKey) {
                return isAtOrBelow(jsonPath, duplicateKeyPath(duplicateKey));
              }
              return problem.jsonPath().map(path -> isAtOrBelow(jsonPath, path)).orElse(false);
            });
  }

  static boolean hasProblemAt(String jsonPath, List<RequestStructuralProblem> problems) {
    return problems.stream()
        .anyMatch(
            problem -> {
              if (problem instanceof RequestDuplicateKey duplicateKey) {
                return duplicateKeyPath(duplicateKey).equals(jsonPath);
              }
              return problem.jsonPath().filter(jsonPath::equals).isPresent();
            });
  }

  private static String duplicateKeyPath(RequestDuplicateKey duplicateKey) {
    // Public duplicate diagnostics intentionally name the containing object only; binding needs
    // the owned child path to prevent an arbitrary first occurrence from becoming usable.
    return RequestObjectMembers.childPath(duplicateKey.containingObjectPath(), duplicateKey.key());
  }

  private static boolean isAtOrBelow(String rootPath, String candidatePath) {
    return candidatePath.equals(rootPath)
        || candidatePath.startsWith(rootPath + ".")
        || candidatePath.startsWith(rootPath + "[");
  }

  private static boolean canConvert(RequestJsonNode node) {
    return switch (node) {
      case RequestJsonInvalid _ -> false;
      case RequestJsonObject object ->
          object.members().stream().allMatch(member -> canConvert(member.value()));
      case RequestJsonArray array ->
          array.elements().stream().allMatch(RequestFragmentBinder::canConvert);
      default -> true;
    };
  }

  static JsonNode toJsonNode(RequestJsonNode node) {
    return switch (node) {
      case RequestJsonObject object -> toObjectNode(object);
      case RequestJsonArray array -> toArrayNode(array);
      case RequestJsonString string -> JsonNodeFactory.instance.stringNode(string.value());
      case RequestJsonNumber number -> numberNode(number.value());
      case RequestJsonBoolean bool -> JsonNodeFactory.instance.booleanNode(bool.value());
      case RequestJsonNull _ -> JsonNodeFactory.instance.nullNode();
      case RequestJsonInvalid invalid ->
          throw new IllegalStateException(
              "Cannot bind invalid JSON node at " + invalid.byteOffset());
    };
  }

  /** Preserves an authored integral token so Jackson does not reject it as a decimal scalar. */
  private static JsonNode numberNode(String value) {
    if (value.indexOf('.') < 0 && value.indexOf('e') < 0 && value.indexOf('E') < 0) {
      return JsonNodeFactory.instance.numberNode(new BigInteger(value));
    }
    return JsonNodeFactory.instance.numberNode(new BigDecimal(value));
  }

  private static ObjectNode toObjectNode(RequestJsonObject object) {
    ObjectNode result = JsonNodeFactory.instance.objectNode();
    for (RequestJsonMember member : object.members()) {
      if (!result.has(member.name())) {
        result.set(member.name(), toJsonNode(member.value()));
      }
    }
    return result;
  }

  private static ArrayNode toArrayNode(RequestJsonArray array) {
    ArrayNode result = JsonNodeFactory.instance.arrayNode();
    for (RequestJsonNode element : array.elements()) {
      result.add(toJsonNode(element));
    }
    return result;
  }
}
