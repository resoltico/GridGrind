package dev.erst.gridgrind.json.codec;

import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import tools.jackson.core.JsonParser;
import tools.jackson.core.JsonToken;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.ValueDeserializer;

/** Resolves grant path literals against the explicit authority-document root. */
final class GrantPathDeserializer extends ValueDeserializer<Path> {
  static final Object ROOT_ATTRIBUTE = GrantPathDeserializer.class;

  @Override
  public Path deserialize(JsonParser parser, DeserializationContext context) {
    if (!parser.hasToken(JsonToken.VALUE_STRING)) {
      return context.reportInputMismatch(Path.class, "Grant paths require a JSON string");
    }
    String value = parser.getString();
    if (value.isBlank()) {
      return context.reportInputMismatch(Path.class, "Grant paths must not be blank");
    }
    Path path;
    try {
      path = Path.of(value);
    } catch (InvalidPathException exception) {
      return context.reportInputMismatch(Path.class, "Grant path is not a valid filesystem path");
    }
    if (path.isAbsolute()) {
      return path;
    }
    @org.jspecify.annotations.Nullable Object root = context.getAttribute(ROOT_ATTRIBUTE);
    if (!(root instanceof Path base)) {
      return context.reportInputMismatch(
          Path.class, "Relative grant paths require an explicit root");
    }
    return base.resolve(path);
  }
}
