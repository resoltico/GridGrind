package dev.erst.gridgrind.contract.json;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadataSupport;
import java.util.Objects;
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
