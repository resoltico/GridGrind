package dev.erst.gridgrind.contract.json;

import org.jspecify.annotations.Nullable;
import tools.jackson.databind.JsonNode;

/** Owns lightweight dotted-and-indexed JSON-path traversal for payload-shape diagnostics. */
final class GridGrindJsonPathSupport {
  private GridGrindJsonPathSupport() {}

  static boolean pathExists(JsonNode root, String jsonPath) {
    JsonNode current = root;
    for (String segment : jsonPath.split("\\.", -1)) {
      current = descendSegment(current, segment);
      if (current == null) {
        return false;
      }
    }
    return true;
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
