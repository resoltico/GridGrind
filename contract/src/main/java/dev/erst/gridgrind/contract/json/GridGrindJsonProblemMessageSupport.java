package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.databind.DatabindException;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.InvalidFormatException;

/** Owns public error wording for JSON failures. */
final class GridGrindJsonProblemMessageSupport {
  private GridGrindJsonProblemMessageSupport() {}

  static IllegalArgumentException invalidPayload(JacksonException exception) {
    GridGrindJsonPayloadMetadataSupport.PayloadMetadata metadata =
        GridGrindJsonPayloadMetadataSupport.payloadMetadata(exception);
    if (exception instanceof StreamReadException) {
      return new InvalidJsonException(
          message(exception),
          metadata.jsonPath(),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    if (exception instanceof DatabindException) {
      return new InvalidRequestShapeException(
          new MessageShape(message(exception), metadata.jsonPath()),
          metadata.jsonPath(),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    return new InvalidJsonException(
        message(exception),
        metadata.jsonPath(),
        metadata.jsonLine(),
        metadata.jsonColumn(),
        exception);
  }

  static IllegalArgumentException invalidPayload(
      JacksonException exception, JsonNode rootNode, Class<?> targetType) {
    Objects.requireNonNull(rootNode, "rootNode must not be null");
    Objects.requireNonNull(targetType, "targetType must not be null");
    GridGrindJsonPayloadMetadataSupport.PayloadMetadata metadata =
        GridGrindJsonPayloadMetadataSupport.payloadMetadata(exception);
    Optional<RequestProblemDescriptor.Shape> structuralProblem =
        GridGrindJsonRequestProblemDetector.detect(rootNode, targetType, exception);
    if (structuralProblem.filter(problem -> !(problem instanceof MessageShape)).isPresent()) {
      RequestProblemDescriptor.Shape requestProblem = structuralProblem.orElseThrow();
      return new InvalidRequestShapeException(
          requestProblem,
          preciseJsonPath(requestProblem, metadata.jsonPath()),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    Optional<Throwable> validationCause = validationCause(exception);
    if (validationCause.isPresent()) {
      return invalidValidationCause(exception, metadata, validationCause.orElseThrow());
    }
    RequestProblemDescriptor.Shape requestProblem =
        structuralProblem.orElseGet(
            () -> new MessageShape(message(exception), metadata.jsonPath()));
    return new InvalidRequestShapeException(
        requestProblem,
        preciseJsonPath(requestProblem, metadata.jsonPath()),
        metadata.jsonLine(),
        metadata.jsonColumn(),
        exception);
  }

  static IllegalArgumentException invalidRequestPayload(JacksonException exception) {
    if (exception instanceof StreamConstraintsException) {
      return GridGrindJsonMapperSupport.requestTooLarge(exception);
    }
    return invalidPayload(exception);
  }

  static String message(Throwable throwable) {
    if (throwable instanceof RequestProblemSource requestProblemSource) {
      return GridGrindRequestProblemSupport.message(requestProblemSource.requestProblem());
    }
    if (throwable instanceof tools.jackson.databind.exc.InvalidTypeIdException invalidTypeId) {
      return GridGrindJsonSubtypeProblemSupport.unknownTypeValueMessage(invalidTypeId);
    }
    if (throwable instanceof tools.jackson.databind.exc.UnrecognizedPropertyException unknown) {
      return "Unknown field '" + unknown.getPropertyName() + "'";
    }
    if (throwable instanceof InvalidFormatException invalidFormat
        && invalidFormat.getTargetType() != null
        && invalidFormat.getTargetType().isEnum()) {
      return GridGrindJsonValueProblemSupport.enumValueMessage(invalidFormat);
    }
    if (throwable instanceof tools.jackson.databind.exc.MismatchedInputException mismatchedInput) {
      return GridGrindJsonValueProblemSupport.mismatchedInputMessage(mismatchedInput);
    }
    String message =
        throwable instanceof JacksonException jacksonException
            ? jacksonException.getOriginalMessage()
            : throwable.getMessage();
    return cleanJacksonMessage(message);
  }

  static String cleanJacksonMessage(@Nullable String message) {
    if (message == null || message.isBlank()) {
      return "Invalid JSON payload";
    }
    String cleaned = message.replaceAll("\\s*\\([^()]*\\[Source:.*\\)$", "").strip();
    return cleaned.isBlank() ? "Invalid JSON payload" : cleaned;
  }

  private static Optional<Throwable> validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof IllegalArgumentException
          || current instanceof java.time.DateTimeException) {
        return Optional.of(current);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static IllegalArgumentException invalidValidationCause(
      JacksonException exception,
      GridGrindJsonPayloadMetadataSupport.PayloadMetadata metadata,
      Throwable cause) {
    Objects.requireNonNull(metadata, "metadata must not be null");
    Objects.requireNonNull(cause, "cause must not be null");
    Optional<PayloadException> payloadCause =
        cause instanceof PayloadException payloadException
            ? Optional.of(payloadException)
            : Optional.empty();
    Optional<String> jsonPath =
        GridGrindJsonPathSupport.qualifyPath(
            metadata.jsonPath(), payloadCause.flatMap(PayloadException::jsonPath));
    Optional<Integer> jsonLine =
        payloadCause.flatMap(PayloadException::jsonLine).or(metadata::jsonLine);
    Optional<Integer> jsonColumn =
        payloadCause.flatMap(PayloadException::jsonColumn).or(metadata::jsonColumn);
    if (cause instanceof InvalidRequestShapeException shapeException) {
      return new InvalidRequestShapeException(
          (RequestProblemDescriptor.Shape) shapeException.requestProblem(),
          jsonPath,
          jsonLine,
          jsonColumn,
          exception);
    }
    if (cause instanceof InvalidRequestException requestException) {
      return new InvalidRequestException(
          (RequestProblemDescriptor.Invariant) requestException.requestProblem(),
          jsonPath,
          jsonLine,
          jsonColumn,
          exception);
    }
    RequestProblemDescriptor.Invariant requestProblem =
        new MessageInvariant(message(cause), jsonPath);
    return new InvalidRequestException(requestProblem, jsonPath, jsonLine, jsonColumn, exception);
  }

  private static Optional<String> preciseJsonPath(
      RequestProblemDescriptor requestProblem, Optional<String> metadataJsonPath) {
    return requestProblem.jsonPath().or(() -> metadataJsonPath);
  }
}
