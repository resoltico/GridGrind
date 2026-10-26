package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.RequestProblemSource;
import dev.erst.gridgrind.core.vocabulary.WireScalar;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.Optional;
import tools.jackson.core.JsonParser;
import tools.jackson.core.JsonToken;
import tools.jackson.databind.DeserializationContext;
import tools.jackson.databind.ValueDeserializer;

/** Binds domain scalars through their validating public constructor without coercing tokens. */
final class WireScalarDeserializer extends ValueDeserializer<WireScalar> {
  private final Constructor<? extends WireScalar> constructor;
  private final Class<?> componentType;

  WireScalarDeserializer(Class<? extends WireScalar> valueType) {
    try {
      componentType = valueType.getRecordComponents()[0].getType();
      constructor = valueType.getConstructor(componentType);
    } catch (NoSuchMethodException exception) {
      throw new IllegalArgumentException(
          "Domain scalars need a public component constructor", exception);
    }
  }

  @Override
  public WireScalar deserialize(JsonParser parser, DeserializationContext context) {
    Object scalar;
    if (componentType == String.class) {
      if (!parser.hasToken(JsonToken.VALUE_STRING)) {
        return context.reportInputMismatch(handledType(), "Domain values require a JSON string");
      }
      scalar = parser.getString();
    } else if (componentType == int.class) {
      if (!parser.hasToken(JsonToken.VALUE_NUMBER_INT)) {
        return context.reportInputMismatch(handledType(), "Domain values require a JSON integer");
      }
      scalar = parser.getIntValue();
    } else {
      throw new IllegalStateException("Unsupported wire scalar component: " + componentType);
    }
    try {
      return constructor.newInstance(scalar);
    } catch (InvocationTargetException exception) {
      Throwable cause = exception.getCause();
      if (cause instanceof IllegalArgumentException invalid
          && !(cause instanceof RequestProblemSource)) {
        String message = invalid.getMessage();
        if (message == null || message.isBlank()) {
          message = "Domain value is invalid.";
        }
        return (WireScalar)
            context.handleInstantiationProblem(
                handledType(),
                scalar,
                new InvalidRequestException(
                    new MessageInvariant(message, Optional.empty()),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    exception));
      }
      return (WireScalar)
          context.handleInstantiationProblem(
              handledType(), scalar, cause == null ? exception : cause);
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException("Unable to construct domain scalar", exception);
    }
  }

  @Override
  public Class<?> handledType() {
    return constructor.getDeclaringClass();
  }
}
