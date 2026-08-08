package dev.erst.gridgrind.contract.json;

import java.util.List;

/** Reads JSON scalar values while retaining syntax diagnostics and input byte offsets. */
final class RequestJsonScalarReader {
  private RequestJsonScalarReader() {}

  static RequestJsonNode read(
      RequestSyntaxCursor cursor, List<RequestStructuralProblem> problems, String jsonPath) {
    int start = cursor.position();
    return switch (cursor.current()) {
      case '"' -> readString(cursor, problems, jsonPath);
      case 't' ->
          readKeyword(
              cursor,
              problems,
              jsonPath,
              "true",
              new RequestJsonBoolean(cursor.byteOffset(), true));
      case 'f' ->
          readKeyword(
              cursor,
              problems,
              jsonPath,
              "false",
              new RequestJsonBoolean(cursor.byteOffset(), false));
      case 'n' ->
          readKeyword(cursor, problems, jsonPath, "null", new RequestJsonNull(cursor.byteOffset()));
      default ->
          isNumberStart(cursor.current())
              ? readNumber(cursor, problems, jsonPath)
              : invalidValue(cursor, problems, jsonPath, start);
    };
  }

  private static RequestJsonString readString(
      RequestSyntaxCursor cursor, List<RequestStructuralProblem> problems, String jsonPath) {
    int start = cursor.position();
    cursor.advance();
    StringBuilder value = new StringBuilder();
    boolean closed = false;
    while (!cursor.atEnd()) {
      char character = cursor.current();
      cursor.advance();
      if (character == '"') {
        closed = true;
        break;
      }
      if (character < 0x20) {
        problem(
            problems,
            "Control character is not allowed in a JSON string",
            cursor.position() - 1,
            cursor,
            jsonPath);
        continue;
      }
      if (character != '\\') {
        value.append(character);
        continue;
      }
      if (cursor.atEnd()) {
        break;
      }
      appendEscapedCharacter(cursor, value, start, problems, jsonPath);
    }
    if (!closed) {
      problem(problems, "Unterminated JSON string", start, cursor, jsonPath);
    }
    return new RequestJsonString(cursor.byteOffsetAt(start), value.toString());
  }

  private static void appendEscapedCharacter(
      RequestSyntaxCursor cursor,
      StringBuilder value,
      int stringStart,
      List<RequestStructuralProblem> problems,
      String jsonPath) {
    char escaped = cursor.current();
    cursor.advance();
    switch (escaped) {
      case '"', '\\', '/' -> value.append(escaped);
      case 'b' -> value.append('\b');
      case 'f' -> value.append('\f');
      case 'n' -> value.append('\n');
      case 'r' -> value.append('\r');
      case 't' -> value.append('\t');
      case 'u' -> appendUnicodeEscape(cursor, value, stringStart, problems, jsonPath);
      default -> {
        problem(problems, "Invalid JSON string escape", cursor.position() - 1, cursor, jsonPath);
        value.append(escaped);
      }
    }
  }

  private static void appendUnicodeEscape(
      RequestSyntaxCursor cursor,
      StringBuilder value,
      int stringStart,
      List<RequestStructuralProblem> problems,
      String jsonPath) {
    if (cursor.position() + 4 > cursor.length()) {
      problem(problems, "Invalid unicode escape in JSON string", stringStart, cursor, jsonPath);
      cursor.moveToEnd();
      return;
    }
    String hexadecimal = cursor.text().substring(cursor.position(), cursor.position() + 4);
    cursor.advanceBy(4);
    try {
      value.append((char) Integer.parseInt(hexadecimal, 16));
    } catch (NumberFormatException exception) {
      problem(problems, "Invalid unicode escape in JSON string", stringStart, cursor, jsonPath);
    }
  }

  private static RequestJsonNode readKeyword(
      RequestSyntaxCursor cursor,
      List<RequestStructuralProblem> problems,
      String jsonPath,
      String keyword,
      RequestJsonNode node) {
    int start = cursor.position();
    if (cursor.text().regionMatches(start, keyword, 0, keyword.length())
        && (start + keyword.length() == cursor.length()
            || isValueDelimiter(cursor.text().charAt(start + keyword.length())))) {
      cursor.advanceBy(keyword.length());
      return node;
    }
    cursor.consumeInvalidValueToken();
    problem(problems, "Invalid JSON literal", start, cursor, jsonPath);
    return new RequestJsonInvalid(cursor.byteOffsetAt(start));
  }

  private static RequestJsonNode readNumber(
      RequestSyntaxCursor cursor, List<RequestStructuralProblem> problems, String jsonPath) {
    int start = cursor.position();
    cursor.consumeInvalidValueToken();
    String value = cursor.text().substring(start, cursor.position());
    if (!value.matches("-?(?:0|[1-9][0-9]*)(?:\\.[0-9]+)?(?:[eE][+-]?[0-9]+)?")) {
      problem(problems, "Invalid JSON number", start, cursor, jsonPath);
      return new RequestJsonInvalid(cursor.byteOffsetAt(start));
    }
    return new RequestJsonNumber(cursor.byteOffsetAt(start), value);
  }

  private static RequestJsonNode invalidValue(
      RequestSyntaxCursor cursor,
      List<RequestStructuralProblem> problems,
      String jsonPath,
      int start) {
    cursor.consumeInvalidValueToken();
    problem(problems, "Expected a JSON value", start, cursor, jsonPath);
    return new RequestJsonInvalid(cursor.byteOffsetAt(start));
  }

  private static void problem(
      List<RequestStructuralProblem> problems,
      String message,
      int characterOffset,
      RequestSyntaxCursor cursor,
      String jsonPath) {
    problems.add(
        new RequestInvalidJson(
            message,
            RequestStructuralProblemSupport.optionalJsonPath(jsonPath),
            java.util.Optional.of(cursor.byteOffsetAt(characterOffset))));
  }

  private static boolean isNumberStart(char character) {
    return "-0123456789".indexOf(character) >= 0;
  }

  private static boolean isValueDelimiter(char character) {
    return RequestSyntaxCursor.isJsonWhitespace(character)
        || character == ','
        || character == ']'
        || character == '}';
  }
}
