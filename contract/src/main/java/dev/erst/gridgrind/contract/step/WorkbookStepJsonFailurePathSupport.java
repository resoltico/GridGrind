package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.json.GridGrindJsonRequestProblemDetector;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.RequestProblemDescriptor;
import dev.erst.gridgrind.contract.json.RequestProblemSource;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JacksonException.Reference;
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
    String nestedFieldPath = nestedFieldPath(node, targetType, failure);
    if (nestedFieldPath.isEmpty()) {
      return fieldName;
    }
    if (nestedFieldPath.equals(fieldName)
        || nestedFieldPath.startsWith(fieldName + ".")
        || nestedFieldPath.startsWith(fieldName + "[")) {
      return nestedFieldPath;
    }
    return nestedFieldPath.startsWith("[")
        ? fieldName + nestedFieldPath
        : fieldName + "." + nestedFieldPath;
  }

  static JacksonException wrapIllegalArgumentFailure(
      JsonParser parser, String fieldName, IllegalArgumentException exception) {
    MismatchedInputException failure =
        MismatchedInputException.from(
            parser,
            WorkbookStep.class,
            Objects.requireNonNullElse(exception.getMessage(), "Invalid request shape"));
    failure.initCause(exception);
    return fieldFailure(qualifiedFieldName(fieldName, null, Object.class, exception), failure);
  }

  static JacksonException wrapJacksonFailure(
      String fieldName, JsonNode node, Class<?> targetType, JacksonException exception) {
    return exception.prependPath(
        WorkbookStep.class,
        exception.getPath().isEmpty()
            ? qualifiedFieldName(fieldName, node, targetType, exception)
            : fieldName);
  }

  private static String nestedFieldPath(
      @Nullable JsonNode node, Class<?> targetType, Exception failure) {
    if (failure instanceof JacksonException jacksonException) {
      String renderedPath = renderPath(jacksonException.getPath());
      if (!renderedPath.isEmpty()) {
        return renderedPath;
      }
      if (node != null) {
        return GridGrindJsonRequestProblemDetector.detect(node, targetType, jacksonException)
            .flatMap(problem -> problem.jsonPath())
            .orElse("");
      }
    }
    if (failure instanceof RequestProblemSource requestProblemSource) {
      Optional<String> jsonPath = requestProblemSource.requestProblem().jsonPath();
      if (jsonPath.isPresent()) {
        return jsonPath.orElseThrow();
      }
    }
    return "";
  }

  private static String renderPath(java.util.List<Reference> path) {
    StringBuilder rendered = new StringBuilder();
    for (Reference reference : path) {
      String propertyName = reference.getPropertyName();
      if (propertyName != null) {
        if (!rendered.isEmpty()) {
          rendered.append('.');
        }
        rendered.append(propertyName);
      } else {
        rendered.append('[').append(reference.getIndex()).append(']');
      }
    }
    return rendered.toString();
  }
}
