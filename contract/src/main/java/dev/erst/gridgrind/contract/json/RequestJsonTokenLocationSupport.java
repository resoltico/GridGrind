package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;

/** Resolves one product-owned JSON path to its authored member or value token. */
final class RequestJsonTokenLocationSupport {
  private RequestJsonTokenLocationSupport() {}

  /**
   * Returns the property-name token for an object field, the value token for an array element, or
   * the root value token for {@code request}.
   */
  static Optional<Long> byteOffsetAt(RequestJsonNode root, String jsonPath) {
    Objects.requireNonNull(root, "root must not be null");
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    if ("request".equals(jsonPath)) {
      return Optional.of(root.byteOffset());
    }
    if (jsonPath.isBlank()) {
      return Optional.empty();
    }
    return locate(root, jsonPath, 0);
  }

  private static Optional<Long> locate(RequestJsonNode node, String jsonPath, int position) {
    if (position >= jsonPath.length()) {
      return Optional.empty();
    }
    return switch (node) {
      case RequestJsonObject object -> locateObjectMember(object, jsonPath, position);
      case RequestJsonArray array -> locateArrayElement(array, jsonPath, position);
      default -> Optional.empty();
    };
  }

  private static Optional<Long> locateObjectMember(
      RequestJsonObject object, String jsonPath, int position) {
    int fieldEnd = nextFieldBoundary(jsonPath, position);
    if (fieldEnd == position) {
      return Optional.empty();
    }
    String fieldName = jsonPath.substring(position, fieldEnd);
    Optional<RequestJsonMember> member =
        object.members().stream()
            .filter(candidate -> candidate.name().equals(fieldName))
            .findFirst();
    if (member.isEmpty()) {
      return Optional.empty();
    }
    RequestJsonMember matched = member.orElseThrow();
    if (fieldEnd == jsonPath.length()) {
      return Optional.of(matched.nameByteOffset());
    }
    int nextPosition = jsonPath.charAt(fieldEnd) == '.' ? fieldEnd + 1 : fieldEnd;
    return locate(matched.value(), jsonPath, nextPosition);
  }

  private static Optional<Long> locateArrayElement(
      RequestJsonArray array, String jsonPath, int position) {
    if (jsonPath.charAt(position) != '[') {
      return Optional.empty();
    }
    int closeBracket = jsonPath.indexOf(']', position);
    if (closeBracket < 0) {
      return Optional.empty();
    }
    Optional<Integer> index = arrayIndex(jsonPath.substring(position + 1, closeBracket));
    if (index.isEmpty() || index.orElseThrow() >= array.elements().size()) {
      return Optional.empty();
    }
    RequestJsonNode element = array.elements().get(index.orElseThrow());
    if (closeBracket + 1 == jsonPath.length()) {
      return Optional.of(element.byteOffset());
    }
    int nextPosition =
        jsonPath.charAt(closeBracket + 1) == '.' ? closeBracket + 2 : closeBracket + 1;
    return locate(element, jsonPath, nextPosition);
  }

  private static int nextFieldBoundary(String jsonPath, int position) {
    int dot = jsonPath.indexOf('.', position);
    int bracket = jsonPath.indexOf('[', position);
    if (dot < 0) {
      return bracket < 0 ? jsonPath.length() : bracket;
    }
    return bracket < 0 ? dot : Math.min(dot, bracket);
  }

  private static Optional<Integer> arrayIndex(String text) {
    try {
      int value = Integer.parseInt(text);
      return value < 0 ? Optional.empty() : Optional.of(value);
    } catch (NumberFormatException ignored) {
      return Optional.empty();
    }
  }
}
