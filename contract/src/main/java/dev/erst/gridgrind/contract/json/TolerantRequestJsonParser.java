package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Parses one UTF-8 request document while retaining every member occurrence and parser fault. */
final class TolerantRequestJsonParser {
  private TolerantRequestJsonParser() {}

  static RequestSyntaxParseResult parse(byte[] bytes) {
    Objects.requireNonNull(bytes, "bytes must not be null");
    RequestUtf8DecodeResult decoded = RequestUtf8Decoder.decode(bytes);
    if (!decoded.problems().isEmpty()) {
      return new RequestSyntaxParseResult(new RequestJsonInvalid(0), decoded.problems());
    }
    if (containsOnlyJsonWhitespace(decoded.text())) {
      return new RequestSyntaxParseResult(
          new RequestJsonInvalid(0),
          List.of(
              new RequestInvalidJson("Invalid JSON payload", Optional.empty(), Optional.of(0L))));
    }
    return new RequestTolerantSyntaxParser(decoded).parseDocument();
  }

  private static boolean containsOnlyJsonWhitespace(String text) {
    return text.isEmpty()
        || text.chars()
            .allMatch(character -> RequestSyntaxCursor.isJsonWhitespace((char) character));
  }
}
