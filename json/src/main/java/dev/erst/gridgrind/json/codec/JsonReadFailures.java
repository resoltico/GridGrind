package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolContractSupport;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.MessageShape;
import dev.erst.gridgrind.core.problem.MissingRequiredField;
import dev.erst.gridgrind.core.problem.MissingTypeDiscriminator;
import dev.erst.gridgrind.core.problem.RequestProblemDescriptor;
import dev.erst.gridgrind.core.problem.RequestProblemSource;
import dev.erst.gridgrind.core.problem.UnknownField;
import dev.erst.gridgrind.core.problem.UnknownTypeValue;
import dev.erst.gridgrind.core.problem.UnsupportedValue;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import tools.jackson.core.JacksonException;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.InvalidFormatException;
import tools.jackson.databind.exc.InvalidTypeIdException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;
import tools.jackson.databind.exc.ValueInstantiationException;

/** Classifies shared wire-read failures from typed data, without reading framework message text. */
final class JsonReadFailures {
  private JsonReadFailures() {}

  static IllegalArgumentException syntax(JacksonException failure) {
    Optional<Integer> line =
        failure.getLocation() != null && failure.getLocation().getLineNr() > 0
            ? Optional.of(failure.getLocation().getLineNr())
            : Optional.empty();
    Optional<Integer> column =
        failure.getLocation() != null && failure.getLocation().getColumnNr() > 0
            ? Optional.of(failure.getLocation().getColumnNr())
            : Optional.empty();
    if (failure instanceof StreamReadException) {
      return new InvalidJsonException(
          "Invalid JSON payload", Optional.empty(), line, column, failure);
    }
    return new InvalidRequestShapeException(
        new MessageShape("JSON value has the wrong shape for this field", path(failure)),
        path(failure),
        line,
        column,
        failure);
  }

  static IllegalArgumentException value(
      JacksonException failure, JsonNode root, Class<?> fallback) {
    Optional<String> path = path(failure);
    if (failure instanceof UnrecognizedPropertyException unknown) {
      String field = unknown.getPropertyName();
      String location = path.orElse(field);
      if (!location.equals(field) && !location.endsWith("." + field))
        location = child(location, field);
      return shape(new UnknownField(location), Optional.of(location), failure);
    }
    if (failure instanceof InvalidTypeIdException subtype) {
      Class<?> base =
          subtype.getBaseType() == null ? fallback : subtype.getBaseType().getRawClass();
      String discriminator = discriminator(base).orElse("type");
      String location = path.orElse("");
      if (!location.equals(discriminator) && !location.endsWith("." + discriminator))
        location = child(location, discriminator);
      if (subtype.getTypeId() == null)
        return shape(new MissingTypeDiscriminator(location), Optional.of(location), failure);
      return shape(
          new UnknownTypeValue(
              subtype.getTypeId(), Optional.of(location), List.of(), Optional.empty()),
          Optional.of(location),
          failure);
    }
    if (failure instanceof InvalidFormatException format
        && format.getTargetType() != null
        && format.getTargetType().isEnum()) {
      List<String> allowed =
          Arrays.stream(format.getTargetType().getEnumConstants())
              .map(value -> ((Enum<?>) value).name())
              .toList();
      return shape(
          new UnsupportedValue(String.valueOf(format.getValue()), path, allowed), path, failure);
    }
    return modelFailure(failure, root, fallback, path);
  }

  private static IllegalArgumentException modelFailure(
      JacksonException failure, JsonNode root, Class<?> fallback, Optional<String> path) {
    Class<?> target = fallback;
    if (failure instanceof ValueInstantiationException creation && creation.getType() != null) {
      target = creation.getType().getRawClass();
    } else if (failure instanceof MismatchedInputException mismatch
        && mismatch.getTargetType() != null) {
      target = mismatch.getTargetType();
    }
    JsonNode container = at(root, failure);
    if (target.isRecord() && container.isObject()) {
      Optional<String> missing =
          GridGrindProtocolContractSupport.requiredFieldNames(target.asSubclass(Record.class))
              .stream()
              .filter(required -> !container.has(required))
              .findFirst();
      if (missing.isPresent()) {
        String location = child(path.orElse(""), missing.orElseThrow());
        return shape(new MissingRequiredField(location), Optional.of(location), failure);
      }
    }
    Optional<Throwable> cause = validationCause(failure);
    if (cause.isPresent()) {
      return validationFailure(failure, cause.orElseThrow(), target, path);
    }
    return shape(
        new MessageShape("JSON value has the wrong shape for this field", path), path, failure);
  }

  private static Optional<Throwable> validationCause(Throwable failure) {
    Throwable cause = failure.getCause();
    while (cause != null) {
      if (cause instanceof RequestProblemSource || cause instanceof IllegalArgumentException) {
        return Optional.of(cause);
      }
      cause = cause.getCause();
    }
    return Optional.empty();
  }

  private static IllegalArgumentException validationFailure(
      JacksonException failure, Throwable cause, Class<?> target, Optional<String> path) {
    if (cause instanceof RequestProblemSource source) {
      return switch (source.requestProblem()) {
        case RequestProblemDescriptor.Shape problem ->
            shape(problem, path.or(problem::jsonPath), failure);
        case RequestProblemDescriptor.Invariant problem ->
            new InvalidRequestException(
                problem, path.or(problem::jsonPath), Optional.empty(), Optional.empty(), failure);
      };
    }
    String message =
        Optional.ofNullable(cause.getMessage())
            .filter(value -> !value.isBlank())
            .orElse("Model value is invalid.");
    message = RequestDiagnosticRedactor.safeValidationFailureMessage(message, target, path);
    return new InvalidRequestException(
        new MessageInvariant(message, path), path, Optional.empty(), Optional.empty(), failure);
  }

  private static IllegalArgumentException shape(
      RequestProblemDescriptor.Shape problem, Optional<String> path, Throwable cause) {
    return new InvalidRequestShapeException(
        problem, path, Optional.empty(), Optional.empty(), cause);
  }

  private static Optional<String> path(JacksonException failure) {
    String value = "";
    for (JacksonException.Reference reference : failure.getPath()) {
      value =
          reference.getPropertyName() != null
              ? child(value, reference.getPropertyName())
              : value + "[" + reference.getIndex() + "]";
    }
    return value.isEmpty() ? Optional.empty() : Optional.of(value);
  }

  private static String child(String parent, String field) {
    return parent.isEmpty() ? field : parent + "." + field;
  }

  private static Optional<String> discriminator(Class<?> type) {
    Union union = type.getAnnotation(Union.class);
    if (union != null && !union.discriminator().isEmpty())
      return Optional.of(union.discriminator());
    for (Class<?> parent : type.getInterfaces()) {
      Optional<String> value = discriminator(parent);
      if (value.isPresent()) return value;
    }
    return Optional.empty();
  }

  private static JsonNode at(JsonNode root, JacksonException failure) {
    JsonNode node = root;
    for (JacksonException.Reference reference : failure.getPath()) {
      node =
          reference.getPropertyName() != null
              ? node.path(reference.getPropertyName())
              : node.path(reference.getIndex());
    }
    return node;
  }
}
