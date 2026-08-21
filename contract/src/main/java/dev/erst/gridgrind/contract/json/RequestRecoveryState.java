package dev.erst.gridgrind.contract.json;

/** Tracks string escaping and nested containers while a tolerant parser seeks a recovery point. */
final class RequestRecoveryState {
  private boolean inString;
  private boolean escaped;
  private int nesting;

  boolean shouldStop(char character, char firstStop, char secondStop) {
    return !inString
        && nesting == 0
        && (isClosingContainer(character) || character == firstStop || character == secondStop);
  }

  void consume(char character) {
    if (inString) {
      consumeStringCharacter(character);
      return;
    }
    switch (character) {
      case '"' -> inString = true;
      case '{', '[' -> nesting++;
      case '}', ']' -> nesting--;
      default -> {
        // Only strings and container delimiters affect recovery state.
      }
    }
  }

  private void consumeStringCharacter(char character) {
    if (escaped) {
      escaped = false;
    } else if (character == '\\') {
      escaped = true;
    } else if (character == '"') {
      inString = false;
    }
  }

  private static boolean isClosingContainer(char character) {
    return character == '}' || character == ']';
  }
}
