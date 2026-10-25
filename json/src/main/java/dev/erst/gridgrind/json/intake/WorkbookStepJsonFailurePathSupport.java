package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.RequestProblemDescriptor;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonParser;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.MismatchedInputException;

/** Derives precise nested request-field paths for step payload failures. */
final class WorkbookStepJsonFailurePathSupport {
  private WorkbookStepJsonFailurePathSupport() {}

  static MismatchedInputException inputMismatch(
      JsonParser parser, RequestProblemDescriptor.Shape requestProblem) {
    InvalidRequestShapeException shapeException =
        new InvalidRequestShapeException(
            requestProblem, Optional.empty(), Optional.empty(), Optional.empty(), null);
    MismatchedInputException failure =
        MismatchedInputException.from(parser, WorkbookStep.class, shapeException.getMessage());
    failure.initCause(shapeException);
    return failure;
  }

  static JacksonException fieldFailure(String fieldName, JacksonException failure) {
    return failure.prependPath(WorkbookStep.class, fieldName);
  }

  static String qualifiedFieldName(
      String fieldName, @Nullable JsonNode node, Class<?> targetType, Exception failure) {
    return WorkbookStepJsonQualifiedFieldSupport.qualifiedFieldName(
        fieldName, node, targetType, failure);
  }

  static InvalidRequestException wrapIllegalArgumentFailure(
      String fieldName, IllegalArgumentException exception) {
    String qualifiedFieldName = qualifiedFieldName(fieldName, null, Object.class, exception);
    return new InvalidRequestException(
        new MessageInvariant(
            Objects.requireNonNullElse(exception.getMessage(), "Invalid request"),
            Optional.of(qualifiedFieldName)),
        Optional.of(qualifiedFieldName),
        Optional.empty(),
        Optional.empty(),
        exception);
  }

  static Optional<InvalidRequestException> wrapValidationJacksonFailure(
      String fieldName, @Nullable JsonNode node, Class<?> targetType, JacksonException exception) {
    Throwable validationCause = validationCause(exception);
    if (validationCause == null) {
      return Optional.empty();
    }
    String qualifiedFieldName = qualifiedFieldName(fieldName, node, targetType, exception);
    if (validationCause instanceof InvalidRequestException requestException) {
      String qualifiedRequestProblemPath =
          qualifiedFieldName(fieldName, null, Object.class, requestException);
      return Optional.of(
          new InvalidRequestException(
              (dev.erst.gridgrind.core.problem.RequestProblemDescriptor.Invariant)
                  requestException.requestProblem(),
              Optional.of(qualifiedRequestProblemPath),
              Optional.empty(),
              Optional.empty(),
              exception));
    }
    return Optional.of(
        new InvalidRequestException(
            new MessageInvariant(
                Objects.requireNonNullElse(validationCause.getMessage(), "Invalid request"),
                Optional.of(qualifiedFieldName)),
            Optional.of(qualifiedFieldName),
            Optional.empty(),
            Optional.empty(),
            exception));
  }

  static JacksonException wrapJacksonFailure(
      String fieldName, JsonNode node, Class<?> targetType, JacksonException exception) {
    return exception.prependPath(
        WorkbookStep.class,
        exception.getPath().isEmpty()
            ? qualifiedFieldName(fieldName, node, targetType, exception)
            : fieldName);
  }

  private static @Nullable Throwable validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof InvalidRequestShapeException) {
        return null;
      }
      if (current instanceof InvalidRequestException
          || current instanceof IllegalArgumentException
          || current instanceof java.time.DateTimeException) {
        return current;
      }
      current = current.getCause();
    }
    return null;
  }
}
