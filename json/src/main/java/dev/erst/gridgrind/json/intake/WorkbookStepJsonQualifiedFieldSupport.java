package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.problem.RequestProblemSource;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JacksonException.Reference;
import tools.jackson.databind.JsonNode;

/** Derives the qualified authored field path for nested step payload failures. */
final class WorkbookStepJsonQualifiedFieldSupport {
  private WorkbookStepJsonQualifiedFieldSupport() {}

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

  private static String nestedFieldPath(
      @Nullable JsonNode node, Class<?> targetType, Exception failure) {
    if (failure instanceof JacksonException jacksonException) {
      String renderedPath = renderPath(jacksonException.getPath());
      if (!renderedPath.isEmpty()) {
        return renderedPath;
      }
      return missingField(node, targetType, jacksonException);
    }
    if (failure instanceof RequestProblemSource requestProblemSource) {
      Optional<String> jsonPath = requestProblemSource.requestProblem().jsonPath();
      if (jsonPath.isPresent()) {
        return jsonPath.orElseThrow();
      }
    }
    return "";
  }

  private static String missingField(
      @Nullable JsonNode node, Class<?> targetType, JacksonException failure) {
    if (node == null || !node.isObject()) {
      return "";
    }
    Class<?> actual = targetType;
    if (failure instanceof tools.jackson.databind.exc.ValueInstantiationException creation
        && creation.getType() != null) {
      actual = creation.getType().getRawClass();
    } else if (failure instanceof tools.jackson.databind.exc.MismatchedInputException mismatch
        && mismatch.getTargetType() != null) {
      actual = mismatch.getTargetType();
    }
    if (!actual.isRecord()) {
      return "";
    }
    return dev.erst.gridgrind.core.catalog.GridGrindProtocolContractSupport.requiredFieldNames(
            actual.asSubclass(Record.class))
        .stream()
        .filter(name -> !node.has(name))
        .findFirst()
        .orElse("");
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
