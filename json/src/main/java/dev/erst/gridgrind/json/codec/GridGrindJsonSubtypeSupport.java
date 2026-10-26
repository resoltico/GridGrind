package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport;
import tools.jackson.databind.jsontype.NamedType;

/** Registers sealed wire alternatives from their core metadata. */
public final class GridGrindJsonSubtypeSupport {
  private GridGrindJsonSubtypeSupport() {}

  static NamedType[] namedLeafSubtypes(Class<?> rootType) {
    return ProtocolTypeMetadataSupport.subtypesForType(rootType).stream()
        .map(type -> new NamedType(type, ProtocolTypeMetadataSupport.requiredTypeId(type)))
        .toArray(NamedType[]::new);
  }

  /** Returns the declared variant identifiers for a wire union. */
  public static java.util.List<String> typeIds(Class<?> rootType) {
    return ProtocolTypeMetadataSupport.subtypesForType(rootType).stream()
        .map(ProtocolTypeMetadataSupport::requiredTypeId)
        .sorted()
        .toList();
  }
}
