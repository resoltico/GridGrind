package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.vocabulary.WireString;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import tools.jackson.core.JsonParser;
import tools.jackson.core.JsonToken;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.ValueDeserializer;

/** Binds domain strings through their validating public constructor without coercing tokens. */
final class WireStringDeserializer extends ValueDeserializer<WireString> {
  private final Constructor<? extends WireString> constructor;

  WireStringDeserializer(Class<? extends WireString> valueType) {
    try {
      constructor = valueType.getConstructor(String.class);
    } catch (NoSuchMethodException exception) {
      throw new IllegalArgumentException(
          "Domain strings need a public String constructor", exception);
    }
  }

  @Override
  public WireString deserialize(JsonParser parser, DeserializationContext context) {
    if (!parser.hasToken(JsonToken.VALUE_STRING)) {
      return context.reportInputMismatch(handledType(), "Domain values require a JSON string");
    }
    String text = parser.getString();
    try {
      return constructor.newInstance(text);
    } catch (InvocationTargetException exception) {
      Throwable cause = exception.getCause();
      return (WireString)
          context.handleInstantiationProblem(
              handledType(), text, cause == null ? exception : cause);
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException("Unable to construct domain string", exception);
    }
  }

  @Override
  public Class<?> handledType() {
    return constructor.getDeclaringClass();
  }
}
