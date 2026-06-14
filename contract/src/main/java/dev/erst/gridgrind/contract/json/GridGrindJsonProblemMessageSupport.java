package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.databind.DatabindException;
import tools.jackson.databind.exc.InvalidFormatException;

/** Owns public error wording for JSON failures. */
final class GridGrindJsonProblemMessageSupport {
  private GridGrindJsonProblemMessageSupport() {}

  static IllegalArgumentException invalidPayload(JacksonException exception) {
    GridGrindJsonPayloadMetadataSupport.PayloadMetadata metadata =
        GridGrindJsonPayloadMetadataSupport.payloadMetadata(exception);
    Optional<Throwable> validationCause = validationCause(exception);
    GridGrindJsonPayloadMetadataSupport.PayloadMetadata effectiveMetadata =
        effectivePayloadMetadata(metadata, validationCause, message(exception));
    if (exception instanceof StreamReadException) {
      return new InvalidJsonException(
          message(exception),
          effectiveMetadata.jsonPath(),
          effectiveMetadata.jsonLine(),
          effectiveMetadata.jsonColumn(),
          exception);
    }
    if (validationCause.isPresent()) {
      Throwable cause = validationCause.orElseThrow();
      String publicMessage = message(cause);
      GridGrindJsonPayloadMetadataSupport.PayloadMetadata validationMetadata =
          effectivePayloadMetadata(metadata, Optional.of(cause), publicMessage);
      if (GridGrindRequestProblemSupport.looksLikeRequestShapeViolation(publicMessage)) {
        return new InvalidRequestShapeException(
            publicMessage,
            validationMetadata.jsonPath(),
            validationMetadata.jsonLine(),
            validationMetadata.jsonColumn(),
            exception);
      }
      return new InvalidRequestException(
          publicMessage,
          validationMetadata.jsonPath(),
          validationMetadata.jsonLine(),
          validationMetadata.jsonColumn(),
          exception);
    }
    if (exception instanceof DatabindException) {
      return new InvalidRequestShapeException(
          message(exception),
          effectiveMetadata.jsonPath(),
          effectiveMetadata.jsonLine(),
          effectiveMetadata.jsonColumn(),
          exception);
    }
    return new InvalidJsonException(
        message(exception),
        effectiveMetadata.jsonPath(),
        effectiveMetadata.jsonLine(),
        effectiveMetadata.jsonColumn(),
        exception);
  }

  static IllegalArgumentException invalidRequestPayload(JacksonException exception) {
    if (exception instanceof StreamConstraintsException) {
      return GridGrindJsonMapperSupport.requestTooLarge(exception);
    }
    return invalidPayload(exception);
  }

  static String message(Throwable throwable) {
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
    return GridGrindJsonValueProblemSupport.productOwnedJacksonMessage(
        cleanJacksonMessage(message));
  }

  static String cleanJacksonMessage(@Nullable String message) {
    if (message == null || message.isBlank()) {
      return "Invalid JSON payload";
    }
    int startMarkerIndex = message.indexOf(" (start marker at [Source:");
    String trimmed = startMarkerIndex >= 0 ? message.substring(0, startMarkerIndex) : message;
    String cleaned =
        trimmed
            .replaceAll(" as a subtype of `[^`]+`", "")
            .replaceAll(" \\(for POJO property '[^']+'\\)", "")
            .replaceAll(" \\(but could if coercion[^)]*\\)", "")
            .strip();
    return cleaned.isBlank() ? "Invalid JSON payload" : cleaned;
  }

  private static Optional<Throwable> validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof IllegalArgumentException
          || current instanceof java.time.DateTimeException) {
        return Optional.of(current);
      }
      if (current instanceof NullPointerException npe && isExplicitNullCheck(npe)) {
        return Optional.of(current);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static boolean isExplicitNullCheck(NullPointerException npe) {
    String message = npe.getMessage();
    return message != null && message.endsWith("must not be null");
  }

  private static GridGrindJsonPayloadMetadataSupport.PayloadMetadata effectivePayloadMetadata(
      GridGrindJsonPayloadMetadataSupport.PayloadMetadata metadata,
      Optional<Throwable> validationCause,
      String publicMessage) {
    Objects.requireNonNull(metadata, "metadata must not be null");
    Objects.requireNonNull(validationCause, "validationCause must not be null");
    Optional<PayloadException> payloadCause =
        validationCause
            .filter(PayloadException.class::isInstance)
            .map(PayloadException.class::cast);
    Optional<String> jsonPath =
        payloadCause
            .flatMap(PayloadException::jsonPath)
            .or(metadata::jsonPath)
            .or(() -> GridGrindRequestProblemSupport.jsonPathFromMessage(publicMessage));
    Optional<Integer> jsonLine =
        payloadCause.flatMap(PayloadException::jsonLine).or(metadata::jsonLine);
    Optional<Integer> jsonColumn =
        payloadCause.flatMap(PayloadException::jsonColumn).or(metadata::jsonColumn);
    return new GridGrindJsonPayloadMetadataSupport.PayloadMetadata(jsonPath, jsonLine, jsonColumn);
  }
}
