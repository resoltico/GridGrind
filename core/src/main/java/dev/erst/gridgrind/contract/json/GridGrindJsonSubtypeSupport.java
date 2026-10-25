package dev.erst.gridgrind.contract.json;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import tools.jackson.databind.jsontype.NamedType;

/** Reflective JSON subtype registration support for sealed GridGrind contract families. */
final class GridGrindJsonSubtypeSupport {
  private GridGrindJsonSubtypeSupport() {}

  static NamedType[] namedLeafSubtypes(Class<?> rootType) {
    return sealedLeafSubtypes(rootType).stream()
        .map(GridGrindJsonSubtypeSupport::namedType)
        .toArray(NamedType[]::new);
  }

  static java.util.List<String> typeIds(Class<?> rootType) {
    com.fasterxml.jackson.annotation.JsonSubTypes jsonSubTypes =
        rootType.getAnnotation(com.fasterxml.jackson.annotation.JsonSubTypes.class);
    if (jsonSubTypes != null) {
      return java.util.Arrays.stream(jsonSubTypes.value())
          .map(com.fasterxml.jackson.annotation.JsonSubTypes.Type::name)
          .filter(name -> !name.isBlank())
          .sorted()
          .toList();
    }
    return sealedLeafSubtypes(rootType).stream()
        .map(GridGrindJsonSubtypeSupport::safeTypeName)
        .flatMap(Optional::stream)
        .sorted()
        .toList();
  }

  private static Optional<String> safeTypeName(Class<?> subtype) {
    JsonTypeName typeName = subtype.getAnnotation(JsonTypeName.class);
    if (typeName != null && !typeName.value().isBlank()) {
      return Optional.of(typeName.value());
    }
    dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata metadata =
        subtype.getAnnotation(dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata.class);
    if (metadata != null && !metadata.id().isBlank()) {
      return Optional.of(metadata.id());
    }
    return Optional.empty();
  }

  private static Set<Class<?>> sealedLeafSubtypes(Class<?> rootType) {
    Objects.requireNonNull(rootType, "rootType must not be null");
    Set<Class<?>> leaves = new java.util.LinkedHashSet<>();
    collectLeafSubtypes(rootType, leaves);
    return Set.copyOf(leaves);
  }

  private static void collectLeafSubtypes(Class<?> classType, Set<Class<?>> leaves) {
    if (classType.isSealed()) {
      for (Class<?> permitted : classType.getPermittedSubclasses()) {
        collectLeafSubtypes(permitted, leaves);
      }
      return;
    }
    leaves.add(classType);
  }

  private static NamedType namedType(Class<?> subtype) {
    JsonTypeName typeName = subtype.getAnnotation(JsonTypeName.class);
    if (typeName != null && !typeName.value().isBlank()) {
      return new NamedType(subtype, typeName.value());
    }
    return new NamedType(subtype, ProtocolTypeMetadataSupport.requiredTypeId(subtype));
  }
}
