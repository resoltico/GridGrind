package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.RequestProblemDescriptor;
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

  static JacksonException wrapJacksonFailure(
      String fieldName, JsonNode node, Class<?> targetType, JacksonException exception) {
    return exception.prependPath(
        WorkbookStep.class,
        exception.getPath().isEmpty()
            ? qualifiedFieldName(fieldName, node, targetType, exception)
            : fieldName);
  }
}
