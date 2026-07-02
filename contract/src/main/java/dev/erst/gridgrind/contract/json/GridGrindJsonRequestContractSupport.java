package dev.erst.gridgrind.contract.json;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import java.lang.reflect.Field;
import java.lang.reflect.RecordComponent;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import tools.jackson.databind.JsonNode;

/** Effective creator/discriminator contract support for request-shape diagnostics. */
final class GridGrindJsonRequestContractSupport {
  private GridGrindJsonRequestContractSupport() {}

  static Optional<RequestProblemDescriptor.Shape> missingRequiredComponentProblem(
      JsonNode rootNode, Class<?> targetType, String basePath) {
    Objects.requireNonNull(rootNode, "rootNode must not be null");
    Objects.requireNonNull(targetType, "targetType must not be null");
    JsonNode containerNode = subtreeAt(rootNode, basePath).orElse(rootNode);
    if (!containerNode.isObject()) {
      return Optional.empty();
    }
    return requiredComponents(targetType).stream()
        .filter(component -> !containerNode.has(component.wireName()))
        .findFirst()
        .map(component -> component.toProblem(basePath));
  }

  static Optional<JsonNode> subtreeAt(JsonNode rootNode, String jsonPath) {
    if (jsonPath.isBlank()) {
      return Optional.of(rootNode);
    }
    return GridGrindJsonPathSupport.nodeAt(rootNode, jsonPath);
  }

  private static List<RequiredComponent> requiredComponents(Class<?> targetType) {
    List<RequiredComponent> requiredComponents = new ArrayList<>();
    GridGrindProtocolContractSupport.discriminatorField(targetType)
        .ifPresent(
            discriminator ->
                requiredComponents.add(
                    new RequiredComponent(discriminator, ComponentKind.DISCRIMINATOR)));
    if (Record.class.isAssignableFrom(targetType)) {
      @SuppressWarnings("unchecked")
      Class<? extends Record> recordType = (Class<? extends Record>) targetType;
      List<String> requiredFieldNames =
          GridGrindProtocolContractSupport.requiredFieldNames(recordType);
      for (RecordComponent component : recordType.getRecordComponents()) {
        if (isIgnored(component)) {
          continue;
        }
        String wireName = wireName(component);
        if (requiredFieldNames.contains(component.getName())) {
          requiredComponents.add(new RequiredComponent(wireName, ComponentKind.FIELD));
        }
      }
    }
    return List.copyOf(requiredComponents);
  }

  private static boolean isIgnored(RecordComponent component) {
    return recordField(component).isAnnotationPresent(JsonIgnore.class)
        || component.getAccessor().isAnnotationPresent(JsonIgnore.class);
  }

  private static String wireName(RecordComponent component) {
    JsonProperty fieldProperty = recordField(component).getAnnotation(JsonProperty.class);
    if (fieldProperty != null && !fieldProperty.value().isBlank()) {
      return fieldProperty.value();
    }
    JsonProperty accessorProperty = component.getAccessor().getAnnotation(JsonProperty.class);
    if (accessorProperty != null && !accessorProperty.value().isBlank()) {
      return accessorProperty.value();
    }
    return component.getName();
  }

  private static Field recordField(RecordComponent component) {
    return Arrays.stream(component.getDeclaringRecord().getDeclaredFields())
        .filter(field -> field.getName().equals(component.getName()))
        .findFirst()
        .orElseThrow();
  }

  private record RequiredComponent(String wireName, ComponentKind kind) {
    private RequiredComponent {
      wireName = RequestProblemDescriptorSupport.requireNonBlank(wireName, "wireName");
      Objects.requireNonNull(kind, "kind must not be null");
    }

    private RequestProblemDescriptor.Shape toProblem(String basePath) {
      String jsonPath =
          GridGrindJsonRequestTypeProblemSupport.pathAlreadyTargetsField(basePath, wireName)
              ? basePath
              : GridGrindJsonRequestTypeProblemSupport.appendPath(basePath, wireName);
      return switch (kind) {
        case FIELD -> new MissingRequiredField(jsonPath);
        case DISCRIMINATOR -> new MissingTypeDiscriminator(jsonPath);
      };
    }
  }

  /** Distinguishes field requirements from discriminator requirements. */
  private enum ComponentKind {
    /** One ordinary creator field declared by the request contract. */
    FIELD,
    /** One discriminator field declared by a typed protocol root. */
    DISCRIMINATOR
  }
}
