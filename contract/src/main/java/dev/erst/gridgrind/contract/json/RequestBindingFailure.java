package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;

/**
 * One constructor-level request failure found while independently binding a valid JSON fragment.
 */
public record RequestBindingFailure(
    IllegalArgumentException exception, String jsonPath, Optional<Long> byteOffset) {
  /** Normalizes one binding failure into the request-problem exception family. */
  static RequestBindingFailure from(
      RuntimeException failure,
      RequestJsonNode fragment,
      String fragmentPath,
      RequestDiagnosticRedactor redactor) {
    Objects.requireNonNull(failure, "failure must not be null");
    Objects.requireNonNull(fragment, "fragment must not be null");
    Objects.requireNonNull(fragmentPath, "fragmentPath must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    NormalizedFailure normalized = normalize(failure, fragment, fragmentPath, redactor);
    String qualifiedPath = normalized.payloadException().jsonPath().orElse(fragmentPath);
    return new RequestBindingFailure(normalized.exception(), qualifiedPath, Optional.empty());
  }

  /** Retains a complete-plan invariant at its already-qualified product-owned request path. */
  static RequestBindingFailure fromCompletePlan(
      RuntimeException failure, RequestDiagnosticRedactor redactor) {
    Objects.requireNonNull(failure, "failure must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    Optional<RequestProblemSource> requestProblem = requestProblem(failure);
    if (requestProblem.isPresent()) {
      return fromCompletePlanRequestProblem(requestProblem.orElseThrow(), failure);
    }
    String jsonPath = "request";
    InvalidRequestException normalized =
        new InvalidRequestException(
            new MessageInvariant(
                redactor.safeBindingFailureMessage(failure.getMessage(), Optional.of(jsonPath)),
                Optional.of(jsonPath)),
            Optional.of(jsonPath),
            Optional.empty(),
            Optional.empty(),
            failure);
    return new RequestBindingFailure(normalized, jsonPath, Optional.empty());
  }

  public RequestBindingFailure {
    Objects.requireNonNull(exception, "exception must not be null");
    jsonPath = requireNonBlank(jsonPath, "jsonPath");
    byteOffset = Objects.requireNonNullElseGet(byteOffset, Optional::empty);
    byteOffset.ifPresent(
        offset -> {
          if (offset < 0) {
            throw new IllegalArgumentException("byteOffset must not be negative");
          }
        });
  }

  /** Returns this failure with its exact request-token location when the raw tree retains one. */
  RequestBindingFailure locatedAt(Optional<Long> byteOffset) {
    return new RequestBindingFailure(exception, jsonPath, byteOffset);
  }

  private static NormalizedFailure normalize(
      RuntimeException failure,
      RequestJsonNode fragment,
      String fragmentPath,
      RequestDiagnosticRedactor redactor) {
    Optional<RequestProblemSource> requestProblem = requestProblem(failure);
    if (requestProblem.isPresent()) {
      return normalizeRequestProblem(requestProblem.orElseThrow(), failure, fragment, fragmentPath);
    }
    Optional<tools.jackson.core.JacksonException> jacksonFailure = jacksonFailure(failure);
    if (jacksonFailure.isPresent()) {
      IllegalArgumentException translated =
          GridGrindJsonProblemMessageSupport.invalidPayload(jacksonFailure.orElseThrow());
      return normalize(translated, fragment, fragmentPath, redactor);
    }
    String message =
        redactor.safeBindingFailureMessage(failure.getMessage(), Optional.of(fragmentPath));
    InvalidRequestException normalized =
        new InvalidRequestException(
            new MessageInvariant(message, Optional.of(fragmentPath)),
            Optional.of(fragmentPath),
            Optional.empty(),
            Optional.empty(),
            failure);
    return new NormalizedFailure(normalized, normalized);
  }

  private static NormalizedFailure normalizeRequestProblem(
      RequestProblemSource failure,
      RuntimeException cause,
      RequestJsonNode fragment,
      String fragmentPath) {
    Optional<String> nestedPath =
        jacksonFailure(cause)
            .flatMap(
                exception ->
                    GridGrindJsonPayloadMetadataSupport.payloadMetadata(exception).jsonPath());
    Optional<String> preciseInnerPath =
        preciseInnerPath(fragment, nestedPath, failure.requestProblem().jsonPath());
    Optional<String> qualifiedPath =
        GridGrindJsonPathSupport.qualifyPath(Optional.of(fragmentPath), preciseInnerPath);
    return switch (failure.requestProblem()) {
      case RequestProblemDescriptor.Shape shape -> {
        InvalidRequestShapeException normalized =
            new InvalidRequestShapeException(
                shape, qualifiedPath, Optional.empty(), Optional.empty(), cause);
        yield new NormalizedFailure(normalized, normalized);
      }
      case RequestProblemDescriptor.Invariant invariant -> {
        InvalidRequestException normalized =
            new InvalidRequestException(
                invariant, qualifiedPath, Optional.empty(), Optional.empty(), cause);
        yield new NormalizedFailure(normalized, normalized);
      }
    };
  }

  private static RequestBindingFailure fromCompletePlanRequestProblem(
      RequestProblemSource failure, RuntimeException cause) {
    String jsonPath = failure.requestProblem().jsonPath().orElse("request");
    return switch (failure.requestProblem()) {
      case RequestProblemDescriptor.Shape shape ->
          new RequestBindingFailure(
              new InvalidRequestShapeException(
                  shape, Optional.of(jsonPath), Optional.empty(), Optional.empty(), cause),
              jsonPath,
              Optional.empty());
      case RequestProblemDescriptor.Invariant invariant ->
          new RequestBindingFailure(
              new InvalidRequestException(
                  invariant, Optional.of(jsonPath), Optional.empty(), Optional.empty(), cause),
              jsonPath,
              Optional.empty());
    };
  }

  /** Keeps a Jackson path only when it names a real node in the independently bound fragment. */
  private static Optional<String> preciseInnerPath(
      RequestJsonNode fragment, Optional<String> jacksonPath, Optional<String> requestProblemPath) {
    if (jacksonPath
        .filter(
            path ->
                GridGrindJsonPathSupport.pathExists(
                    RequestFragmentBinder.toJsonNode(fragment), path))
        .isPresent()) {
      return GridGrindJsonPathSupport.qualifyPath(jacksonPath, requestProblemPath);
    }
    return requestProblemPath.or(() -> jacksonPath);
  }

  private static Optional<RequestProblemSource> requestProblem(Throwable failure) {
    Throwable current = failure;
    while (current != null) {
      if (current instanceof RequestProblemSource requestProblemSource) {
        return Optional.of(requestProblemSource);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static Optional<tools.jackson.core.JacksonException> jacksonFailure(Throwable failure) {
    Throwable current = failure;
    while (current != null) {
      if (current instanceof tools.jackson.core.JacksonException jacksonException) {
        return Optional.of(jacksonException);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private record NormalizedFailure(
      IllegalArgumentException exception, PayloadException payloadException) {
    private NormalizedFailure {
      Objects.requireNonNull(exception, "exception must not be null");
      Objects.requireNonNull(payloadException, "payloadException must not be null");
    }
  }
}
