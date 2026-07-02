package dev.erst.gridgrind.contract.json;

import org.jspecify.annotations.Nullable;
import tools.jackson.databind.JsonNode;

/** Owns lightweight dotted-and-indexed JSON-path traversal for payload-shape diagnostics. */
final class GridGrindJsonPathSupport {
  private GridGrindJsonPathSupport() {}

  static java.util.Optional<String> qualifyPath(
      java.util.Optional<String> outerPath, java.util.Optional<String> innerPath) {
    if (innerPath.isEmpty()) {
      return outerPath;
    }
    if (outerPath.isEmpty()) {
      return innerPath;
    }
    String outer = outerPath.orElseThrow();
    String inner = innerPath.orElseThrow();
    if (inner.equals(outer) || inner.startsWith(outer + ".") || inner.startsWith(outer + "[")) {
      return innerPath;
    }
    if (outer.endsWith("." + inner) || (inner.startsWith("[") && outer.endsWith(inner))) {
      return outerPath;
    }
    return java.util.Optional.of(inner.startsWith("[") ? outer + inner : outer + "." + inner);
  }

  static boolean pathExists(JsonNode root, String jsonPath) {
    return nodeAt(root, jsonPath).isPresent();
  }

  static java.util.Optional<JsonNode> nodeAt(JsonNode root, String jsonPath) {
    JsonNode current = root;
    for (String segment : jsonPath.split("\\.", -1)) {
      current = descendSegment(current, segment);
      if (current == null) {
        return java.util.Optional.empty();
      }
    }
    return java.util.Optional.of(current);
  }

  private static @Nullable JsonNode descendSegment(JsonNode current, String segment) {
    JsonNode cursor = current;
    int bracketStart = segment.indexOf('[');
    String propertyName = bracketStart >= 0 ? segment.substring(0, bracketStart) : segment;
    if (!propertyName.isEmpty()) {
      cursor = cursor.get(propertyName);
      if (cursor == null) {
        return null;
      }
    }
    int index = bracketStart;
    while (index >= 0) {
      int closingBracket = segment.indexOf(']', index);
      if (closingBracket < 0) {
        return null;
      }
      int arrayIndex = Integer.parseInt(segment.substring(index + 1, closingBracket));
      cursor = cursor.get(arrayIndex);
      if (cursor == null) {
        return null;
      }
      index = segment.indexOf('[', closingBracket + 1);
    }
    return cursor;
  }
}
