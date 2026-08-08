package dev.erst.gridgrind.contract.json;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Builds a recoverable request syntax tree while preserving duplicate object members. */
final class RequestTolerantSyntaxParser {
  private final RequestSyntaxCursor cursor;
  private final List<RequestStructuralProblem> problems = new ArrayList<>();

  RequestTolerantSyntaxParser(RequestUtf8DecodeResult input) {
    cursor = new RequestSyntaxCursor(input);
  }

  RequestSyntaxParseResult parseDocument() {
    cursor.skipWhitespace();
    RequestJsonNode root = parseValue("");
    cursor.skipWhitespace();
    if (!cursor.atEnd()) {
      problem("Unexpected token after the root JSON value", cursor.position(), "");
      cursor.moveToEnd();
    }
    return new RequestSyntaxParseResult(root, problems);
  }

  private RequestJsonNode parseValue(String jsonPath) {
    cursor.skipWhitespace();
    if (cursor.atEnd()) {
      problem("Expected a JSON value", cursor.position(), jsonPath);
      return new RequestJsonInvalid(cursor.byteOffset());
    }
    return switch (cursor.current()) {
      case '{' -> parseObject(jsonPath);
      case '[' -> parseArray(jsonPath);
      default -> RequestJsonScalarReader.read(cursor, problems, jsonPath);
    };
  }

  @SuppressWarnings(
      "PMD.UseConcurrentHashMap") // Thread-confined insertion order makes duplicate reports stable.
  private RequestJsonNode parseObject(String jsonPath) {
    int start = cursor.position();
    cursor.advance();
    List<RequestJsonMember> members = new ArrayList<>();
    Map<String, Integer> duplicateCounts = new LinkedHashMap<>();
    cursor.skipWhitespace();
    while (!cursor.atEnd() && !atContainerCloser()) {
      if (cursor.current() != '"') {
        problem("Expected an object property name", cursor.position(), jsonPath);
        cursor.recoverObjectMember();
        cursor.skipWhitespace();
        if (cursor.consume(',')) {
          cursor.skipWhitespace();
          continue;
        }
        break;
      }
      int nameOffset = cursor.position();
      RequestJsonString name =
          (RequestJsonString) RequestJsonScalarReader.read(cursor, problems, jsonPath);
      cursor.skipWhitespace();
      if (!cursor.consume(':')) {
        problem(
            "Expected ':' after object property name",
            nameOffset,
            appendPath(jsonPath, name.value()));
        cursor.recoverToValue();
      }
      RequestJsonNode value = parseValue(appendPath(jsonPath, name.value()));
      int priorCount = duplicateCounts.getOrDefault(name.value(), 0);
      duplicateCounts.put(name.value(), priorCount + 1);
      if (priorCount > 0) {
        problems.add(
            new RequestDuplicateKey(
                jsonPath, name.value(), priorCount - 1, cursor.byteOffsetAt(nameOffset)));
      }
      members.add(new RequestJsonMember(name.value(), cursor.byteOffsetAt(nameOffset), value));
      cursor.skipWhitespace();
      if (continuesAfterComma(jsonPath, '}', "object")) {
        continue;
      }
      if (cursor.currentOrZero() != '}') {
        problem("Expected ',' or '}' after object member", cursor.position(), jsonPath);
        cursor.recoverObjectMember();
        cursor.skipWhitespace();
        cursor.consume(',');
        cursor.skipWhitespace();
      }
    }
    if (!cursor.consume('}')) {
      problem("Expected '}' to close object", start, jsonPath);
    }
    return new RequestJsonObject(cursor.byteOffsetAt(start), members);
  }

  private RequestJsonNode parseArray(String jsonPath) {
    int start = cursor.position();
    cursor.advance();
    List<RequestJsonNode> elements = new ArrayList<>();
    cursor.skipWhitespace();
    while (!cursor.atEnd() && !atContainerCloser()) {
      elements.add(parseValue(jsonPath + "[" + elements.size() + "]"));
      cursor.skipWhitespace();
      if (continuesAfterComma(jsonPath, ']', "array")) {
        continue;
      }
      if (cursor.currentOrZero() != ']') {
        problem("Expected ',' or ']' after array element", cursor.position(), jsonPath);
        cursor.recoverArrayElement();
        cursor.skipWhitespace();
        cursor.consume(',');
        cursor.skipWhitespace();
      }
    }
    if (!cursor.consume(']')) {
      problem("Expected ']' to close array", start, jsonPath);
    }
    return new RequestJsonArray(cursor.byteOffsetAt(start), elements);
  }

  private boolean continuesAfterComma(String jsonPath, char closer, String containerName) {
    int separatorOffset = cursor.position();
    if (!cursor.consume(',')) {
      return false;
    }
    cursor.skipWhitespace();
    if (cursor.currentOrZero() != closer) {
      return true;
    }
    problem(
        "Trailing comma is not permitted in a JSON " + containerName, separatorOffset, jsonPath);
    return false;
  }

  private void problem(String message, int characterOffset, String affectedJsonPath) {
    problems.add(
        new RequestInvalidJson(
            message,
            RequestStructuralProblemSupport.optionalJsonPath(affectedJsonPath),
            java.util.Optional.of(cursor.byteOffsetAt(characterOffset))));
  }

  /**
   * A mismatched closer belongs to an enclosing parser frame. Stop this frame without consuming it
   * so its owner can report the missing closer and continue from a stable cursor position.
   */
  private boolean atContainerCloser() {
    return cursor.current() == '}' || cursor.current() == ']';
  }

  private static String appendPath(String basePath, String fieldName) {
    return basePath.isEmpty() ? fieldName : basePath + "." + fieldName;
  }
}
