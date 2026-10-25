package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;

/** Normalizes independently bound request-fragment failures without losing typed causes. */
final class RequestBindingFailureNormalizer {
  private RequestBindingFailureNormalizer() {}

  static RequestBindingFailure from(
      RuntimeException failure,
      RequestJsonNode fragment,
      String fragmentPath,
      RequestDiagnosticRedactor redactor) {
    Optional<FormulaRequestException> formulaProblem =
        FormulaRequestProblemSupport.inputFailure(
            failure,
            RequestBindingPathSupport.qualifiedFormulaFailurePath(fragment, fragmentPath, failure),
            Optional.empty(),
            Optional.empty(),
            failure);
    if (formulaProblem.isPresent()) {
      FormulaRequestException normalized = formulaProblem.orElseThrow();
      Optional<String> inferredPath =
          RequestBindingPathSupport.directChildPathForInvariant(fragment, failure);
      if (inferredPath.isPresent()) {
        String qualifiedPath =
            GridGrindJsonPathSupport.qualifyPath(Optional.of(fragmentPath), inferredPath)
                .orElse(fragmentPath);
        FormulaRequestException rebased =
            new FormulaRequestException(
                normalized.problemCode(),
                normalized.getMessage(),
                Optional.of(qualifiedPath),
                Optional.empty(),
                Optional.empty(),
                failure);
        return new RequestBindingFailure(rebased, qualifiedPath, Optional.empty());
      }
      return new RequestBindingFailure(
          normalized, normalized.jsonPath().orElse(fragmentPath), Optional.empty());
    }
    NormalizedFailure normalized = normalize(failure, fragment, fragmentPath, redactor);
    String qualifiedPath = normalized.payloadException().jsonPath().orElse(fragmentPath);
    Optional<String> inferredPath =
        RequestBindingPathSupport.directChildPathForInvariant(fragment, failure);
    if (inferredPath.isPresent()) {
      String inferredQualifiedPath =
          GridGrindJsonPathSupport.qualifyPath(Optional.of(fragmentPath), inferredPath)
              .orElse(fragmentPath);
      return new RequestBindingFailure(
          RequestBindingFailureRebaser.rebase(
              normalized.exception(), inferredQualifiedPath, failure),
          inferredQualifiedPath,
          Optional.empty());
    }
    return new RequestBindingFailure(normalized.exception(), qualifiedPath, Optional.empty());
  }

  static RequestBindingFailure fromCompletePlan(
      RuntimeException failure, RequestDiagnosticRedactor redactor) {
    Optional<FormulaRequestException> formulaProblem =
        FormulaRequestProblemSupport.inputFailure(
            failure, Optional.of("request"), Optional.empty(), Optional.empty(), failure);
    if (formulaProblem.isPresent()) {
      return new RequestBindingFailure(formulaProblem.orElseThrow(), "request", Optional.empty());
    }
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

  private static NormalizedFailure normalize(
      RuntimeException failure,
      RequestJsonNode fragment,
      String fragmentPath,
      RequestDiagnosticRedactor redactor) {
    Optional<FormulaRequestException> formulaProblem =
        FormulaRequestProblemSupport.requestFailure(failure);
    if (formulaProblem.isPresent()) {
      FormulaRequestException normalized =
          FormulaRequestProblemSupport.normalizeRequestFailure(
              formulaProblem.orElseThrow(), failure, fragment, fragmentPath);
      return new NormalizedFailure(normalized, normalized);
    }
    Optional<RequestProblemSource> requestProblem = requestProblem(failure);
    if (requestProblem.isPresent()) {
      return normalizeRequestProblem(requestProblem.orElseThrow(), failure, fragment, fragmentPath);
    }
    Optional<tools.jackson.core.JacksonException> jacksonFailure =
        RequestBindingPathSupport.jacksonFailure(failure);
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
    Optional<String> preciseInnerPath =
        RequestBindingPathSupport.directChildPathForInvariant(fragment, cause)
            .or(
                () ->
                    RequestBindingPathSupport.preciseInnerPath(
                        fragment,
                        RequestBindingPathSupport.jacksonFailure(cause)
                            .flatMap(
                                exception ->
                                    GridGrindJsonPayloadMetadataSupport.payloadMetadata(exception)
                                        .jsonPath()),
                        failure.requestProblem().jsonPath()));
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

  private record NormalizedFailure(
      IllegalArgumentException exception, PayloadException payloadException) {
    private NormalizedFailure {
      Objects.requireNonNull(exception, "exception must not be null");
      Objects.requireNonNull(payloadException, "payloadException must not be null");
    }
  }
}
