package dev.erst.gridgrind.contract.json;

/** Mutable cursor over one decoded request document with JSON-aware recovery helpers. */
final class RequestSyntaxCursor {
  private final RequestUtf8DecodeResult input;
  private final String text;
  private int index;

  RequestSyntaxCursor(RequestUtf8DecodeResult input) {
    this.input = input;
    this.text = input.text();
  }

  int position() {
    return index;
  }

  int length() {
    return text.length();
  }

  String text() {
    return text;
  }

  long byteOffset() {
    return input.byteOffsetAt(index);
  }

  long byteOffsetAt(int characterOffset) {
    return input.byteOffsetAt(characterOffset);
  }

  boolean atEnd() {
    return index >= text.length();
  }

  char current() {
    return text.charAt(index);
  }

  char currentOrZero() {
    return atEnd() ? '\0' : current();
  }

  void advance() {
    if (!atEnd()) {
      index++;
    }
  }

  boolean consume(char expected) {
    if (currentOrZero() != expected) {
      return false;
    }
    advance();
    return true;
  }

  void skipWhitespace() {
    while (!atEnd() && isJsonWhitespace(current())) {
      advance();
    }
  }

  /** Returns whether one code unit is one of JSON's four grammar-defined whitespace characters. */
  static boolean isJsonWhitespace(char character) {
    return character == ' ' || character == '\n' || character == '\r' || character == '\t';
  }

  void advanceBy(int characterCount) {
    index = Math.min(index + characterCount, text.length());
  }

  void moveToEnd() {
    index = text.length();
  }

  void recoverObjectMember() {
    recoverUntil(',', '}');
  }

  void recoverArrayElement() {
    recoverUntil(',', ']');
  }

  void recoverToValue() {
    skipWhitespace();
    consume(':');
    skipWhitespace();
  }

  void consumeInvalidValueToken() {
    while (!atEnd() && !isValueDelimiter(current())) {
      advance();
    }
  }

  private void recoverUntil(char firstStop, char secondStop) {
    RequestRecoveryState state = new RequestRecoveryState();
    while (!atEnd()) {
      char character = current();
      if (state.shouldStop(character, firstStop, secondStop)) {
        return;
      }
      state.consume(character);
      advance();
    }
  }

  private static boolean isValueDelimiter(char character) {
    return isJsonWhitespace(character) || character == ',' || character == ']' || character == '}';
  }
}
