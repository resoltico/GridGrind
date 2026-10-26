package dev.erst.gridgrind.json.intake;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Resolves one product-owned JSON path to its authored member or value token. */
final class RequestJsonTokenLocationSupport {
  private RequestJsonTokenLocationSupport() {}

  // The mutable index is confined to this call and frozen before leaving the parser boundary.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Map<String, Long> index(RequestJsonNode root) {
    Map<String, Long> offsets = new LinkedHashMap<>();
    offsets.put("request", root.byteOffset());
    collectOffsets(root, "", offsets);
    return Map.copyOf(offsets);
  }

  private static void collectOffsets(RequestJsonNode node, String path, Map<String, Long> offsets) {
    switch (node) {
      case RequestJsonObject object -> {
        Set<String> seen = new HashSet<>();
        for (RequestJsonMember member : object.members()) {
          String name = member.name();
          if (!seen.add(name)) {
            continue;
          }
          String fieldPath = RequestJsonPaths.field(path, name);
          offsets.putIfAbsent(fieldPath, member.nameByteOffset());
          collectOffsets(member.value(), fieldPath, offsets);
        }
      }
      case RequestJsonArray array -> {
        for (int index = 0; index < array.elements().size(); index++) {
          RequestJsonNode element = array.elements().get(index);
          String elementPath = path + "[" + index + "]";
          offsets.putIfAbsent(elementPath, element.byteOffset());
          collectOffsets(element, elementPath, offsets);
        }
      }
      default -> {}
    }
  }

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
    return locate(root, jsonPath, "");
  }

  private static Optional<Long> locate(RequestJsonNode node, String jsonPath, String prefix) {
    if (jsonPath.length() <= prefix.length()) {
      return Optional.empty();
    }
    return switch (node) {
      case RequestJsonObject object -> locateObjectMember(object, jsonPath, prefix);
      case RequestJsonArray array -> locateArrayElement(array, jsonPath, prefix);
      default -> Optional.empty();
    };
  }

  private static Optional<Long> locateObjectMember(
      RequestJsonObject object, String jsonPath, String prefix) {
    for (RequestJsonMember member : object.members()) {
      String fieldPath = RequestJsonPaths.field(prefix, member.name());
      if (jsonPath.equals(fieldPath)) {
        return Optional.of(member.nameByteOffset());
      }
      if (isDescendant(fieldPath, jsonPath)) {
        return locate(member.value(), jsonPath, fieldPath);
      }
    }
    return Optional.empty();
  }

  private static Optional<Long> locateArrayElement(
      RequestJsonArray array, String jsonPath, String prefix) {
    int position = prefix.length();
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
    String elementPath = prefix + "[" + index.orElseThrow() + "]";
    if (jsonPath.equals(elementPath)) {
      return Optional.of(element.byteOffset());
    }
    return isDescendant(elementPath, jsonPath)
        ? locate(element, jsonPath, elementPath)
        : Optional.empty();
  }

  private static boolean isDescendant(String prefix, String path) {
    return path.startsWith(prefix + ".") || path.startsWith(prefix + "[");
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
