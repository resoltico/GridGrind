package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import java.util.Arrays;
import java.util.Optional;
import tools.jackson.core.JacksonException;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.exc.InvalidFormatException;
import tools.jackson.databind.exc.InvalidTypeIdException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;

/** Type and path handling for structural request-shape detection. */
final class GridGrindJsonRequestTypeProblemSupport {
  private GridGrindJsonRequestTypeProblemSupport() {}

  static RequestProblemDescriptor.Shape typeProblem(
      JsonNode rootNode, InvalidTypeIdException exception) {
    String renderedPath = renderedPath(exception);
    String discriminatorField = discriminatorField(exception);
    String fullPath = discriminatorPath(renderedPath, discriminatorField);
    String containerPath = discriminatorContainerPath(renderedPath, discriminatorField);
    JsonNode containerNode =
        GridGrindJsonRequestContractSupport.subtreeAt(rootNode, containerPath).orElse(rootNode);
    String typeId = exception.getTypeId();
    if (typeId != null) {
      JsonNode discriminatorNode =
          containerNode.isObject() ? containerNode.get(discriminatorField) : null;
      if (discriminatorNode != null && !discriminatorNode.isTextual()) {
        return new ActionableShapeMessage(
            "Field '%s' must be a string".formatted(discriminatorField),
            "Replace field '%s' with a JSON string type id.".formatted(discriminatorField),
            Optional.of(fullPath));
      }
      return new UnknownTypeValue(
          typeId,
          Optional.of(fullPath),
          GridGrindJsonSubtypeProblemSupport.similarTypeIds(exception, typeId),
          GridGrindJsonSubtypeProblemSupport.specificGuidance(exception, typeId));
    }
    if (containerNode.isObject() && !containerNode.has(discriminatorField)) {
      return new MissingTypeDiscriminator(fullPath);
    }
    return new MessageShape(
        "JSON object is missing required fields or has the wrong shape",
        optionalPath(containerPath.isBlank() ? renderedPath : containerPath));
  }

  static RequestProblemDescriptor.Shape enumValueProblem(InvalidFormatException exception) {
    Object rawValue = exception.getValue();
    String value = rawValue == null ? "null" : rawValue.toString();
    String jsonPath = renderedPath(exception);
    String[] enumConstants =
        Arrays.stream(exception.getTargetType().getEnumConstants())
            .map(constant -> ((Enum<?>) constant).name())
            .toArray(String[]::new);
    return new UnsupportedValue(value, optionalPath(jsonPath), java.util.List.of(enumConstants));
  }

  static String fullUnknownFieldPath(UnrecognizedPropertyException exception) {
    String renderedPath = renderedPath(exception);
    String propertyName = exception.getPropertyName();
    if (renderedPath.isBlank()) {
      return propertyName;
    }
    if (pathAlreadyTargetsField(renderedPath, propertyName)) {
      return renderedPath;
    }
    return appendPath(renderedPath, propertyName);
  }

  static Optional<String> payloadPath(JacksonException exception) {
    return optionalPath(renderedPath(exception));
  }

  static Optional<String> optionalPath(String jsonPath) {
    return jsonPath.isBlank() ? Optional.empty() : Optional.of(jsonPath);
  }

  static String renderedPath(JacksonException exception) {
    return GridGrindJsonPayloadMetadataSupport.renderPath(exception.getPath());
  }

  static String appendPath(String basePath, String fieldName) {
    if (basePath.isBlank()) {
      return fieldName;
    }
    if (fieldName.startsWith("[")) {
      return basePath + fieldName;
    }
    return basePath + "." + fieldName;
  }

  static boolean pathAlreadyTargetsField(String jsonPath, String fieldName) {
    return jsonPath.equals(fieldName) || jsonPath.endsWith("." + fieldName);
  }

  private static String discriminatorPath(String basePath, String discriminatorField) {
    return pathAlreadyTargetsField(basePath, discriminatorField)
        ? basePath
        : appendPath(basePath, discriminatorField);
  }

  private static String discriminatorContainerPath(String basePath, String discriminatorField) {
    if (basePath.equals(discriminatorField)) {
      return "";
    }
    String fieldSuffix = "." + discriminatorField;
    return basePath.endsWith(fieldSuffix)
        ? basePath.substring(0, basePath.length() - fieldSuffix.length())
        : basePath;
  }

  private static String discriminatorField(InvalidTypeIdException exception) {
    if (exception.getBaseType() == null) {
      return "type";
    }
    return GridGrindProtocolContractSupport.discriminatorField(
            exception.getBaseType().getRawClass())
        .orElse("type");
  }
}
