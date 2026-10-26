package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.semantics.ProtocolTargetingMode;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Reflective support for colocated protocol leaf metadata. */
public final class ProtocolTypeMetadataSupport {
  private ProtocolTypeMetadataSupport() {}

  /** Returns the stable wire id for one protocol leaf type. */
  public static String requiredTypeId(Class<?> type) {
    Objects.requireNonNull(type, "type must not be null");
    Variant metadata = type.getAnnotation(Variant.class);
    if (metadata != null && !metadata.id().isBlank()) {
      return metadata.id();
    }
    throw new IllegalStateException(
        "Contract subtype " + type.getName() + " must declare a non-blank @Variant id");
  }

  /** Returns the colocated metadata for one protocol leaf type. */
  public static Variant requiredMetadata(Class<?> type) {
    Objects.requireNonNull(type, "type must not be null");
    Variant metadata = type.getAnnotation(Variant.class);
    if (metadata == null) {
      throw new IllegalStateException(
          "Contract subtype " + type.getName() + " must declare @Variant");
    }
    if (metadata.id().isBlank()) {
      throw new IllegalStateException(
          "Contract subtype " + type.getName() + " must declare a non-blank metadata id");
    }
    if (metadata.summary().isBlank()) {
      throw new IllegalStateException(
          "Contract subtype " + type.getName() + " must declare a non-blank metadata summary");
    }
    String wireId = requiredTypeId(type);
    if (!wireId.equals(metadata.id())) {
      throw new IllegalStateException(
          "Contract subtype "
              + type.getName()
              + " declares mismatched ids: wire="
              + wireId
              + ", metadata="
              + metadata.id());
    }
    return metadata;
  }

  /**
   * Returns the leaf record subtype owned by one discriminator value, when the family declares it.
   */
  public static Optional<Class<? extends Record>> subtypeForTypeId(
      Class<?> rootType, String typeId) {
    Objects.requireNonNull(rootType, "rootType must not be null");
    Objects.requireNonNull(typeId, "typeId must not be null");
    return typeIdsByClass(rootType).entrySet().stream()
        .filter(entry -> entry.getValue().equals(typeId))
        .<Class<? extends Record>>map(entry -> entry.getKey().asSubclass(Record.class))
        .findFirst();
  }

  /** Returns every concrete record subtype accepted by one discriminated request family. */
  public static List<Class<? extends Record>> subtypesForType(Class<?> rootType) {
    Objects.requireNonNull(rootType, "rootType must not be null");
    return typeIdsByClass(rootType).entrySet().stream()
        .sorted(Map.Entry.comparingByValue())
        .<Class<? extends Record>>map(entry -> entry.getKey().asSubclass(Record.class))
        .toList();
  }

  /** Returns concrete alternatives in their sealed declaration order. */
  public static List<Class<? extends Record>> declaredVariants(Class<?> rootType) {
    Objects.requireNonNull(rootType, "rootType must not be null");
    if (!rootType.isSealed()) {
      throw new IllegalArgumentException(rootType + " must be sealed for discriminator discovery");
    }
    return sealedLeafSubtypes(rootType).stream()
        .<Class<? extends Record>>map(type -> type.asSubclass(Record.class))
        .toList();
  }

  /** Returns the targeting mode declared for one protocol leaf. */
  public static ProtocolTargetingMode targetingMode(Class<?> type) {
    return requiredMetadata(type).targetingMode();
  }

  /** Returns the discovery/help selector rule declared for one protocol leaf. */
  public static Optional<String> targetSelectorRule(Class<?> type) {
    String rule = requiredMetadata(type).targetSelectorRule();
    return rule.isBlank() ? Optional.empty() : Optional.of(rule);
  }

  /** Returns the static selector-family set declared for one protocol leaf. */
  public static Class<? extends Selector>[] staticTargetSelectors(Class<?> type) {
    Variant metadata = requiredMetadata(type);
    if (metadata.targetingMode() != ProtocolTargetingMode.STATIC) {
      throw new IllegalArgumentException(
          "Protocol subtype "
              + type.getName()
              + " derives target selectors dynamically: "
              + targetSelectorRule(type).orElse("no rule declared"));
    }
    return metadata.targetSelectors();
  }

  static Map<Class<?>, String> typeIdsByClass(Class<?> rootType) {
    Objects.requireNonNull(rootType, "rootType must not be null");
    if (!rootType.isSealed()) {
      throw new IllegalArgumentException(rootType + " must be sealed for discriminator discovery");
    }
    return CatalogDescriptorMaps.uniqueMap(
        sealedLeafSubtypes(rootType),
        leaf -> leaf,
        ProtocolTypeMetadataSupport::requiredTypeId,
        "duplicate subtype id ");
  }

  static List<CatalogTypeDescriptor> catalogDescriptorsFor(Class<?> rootType) {
    if (!rootType.isSealed()) {
      throw new IllegalArgumentException(rootType + " must be sealed");
    }
    return sealedLeafSubtypes(rootType).stream()
        .map(ProtocolTypeMetadataSupport::catalogDescriptorFor)
        .toList();
  }

  private static CatalogTypeDescriptor catalogDescriptorFor(Class<?> leafType) {
    return CatalogTypes.descriptor(leafType);
  }

  private static List<Class<?>> sealedLeafSubtypes(Class<?> rootType) {
    Set<Class<?>> leaves = new LinkedHashSet<>();
    collectLeafSubtypes(rootType, leaves);
    return List.copyOf(leaves);
  }

  private static void collectLeafSubtypes(Class<?> type, Set<Class<?>> leaves) {
    if (type.isSealed()) {
      for (Class<?> permitted : type.getPermittedSubclasses()) {
        collectLeafSubtypes(permitted, leaves);
      }
      return;
    }
    leaves.add(type);
  }
}
