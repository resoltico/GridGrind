package dev.erst.gridgrind.core.metadata;

import dev.erst.gridgrind.core.vocabulary.WireScalar;
import java.lang.reflect.Modifier;
import java.lang.reflect.RecordComponent;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/** Resolves declared component defaults without a JSON-library dependency. */
public final class ModelDefaults {
  private ModelDefaults() {}

  /** Returns the declared Java default using already-resolved sibling component values. */
  @SuppressWarnings({"unchecked", "rawtypes"})
  public static Object value(RecordComponent component, Map<String, Object> siblings) {
    WireDefault metadata = component.getAnnotation(WireDefault.class);
    if (metadata == null) {
      throw new IllegalArgumentException("Component has no wire default: " + component);
    }
    if (!metadata.from().isEmpty()) {
      if (!siblings.containsKey(metadata.from())) {
        throw new IllegalStateException(
            "Default depends on unavailable component: " + metadata.from());
      }
      return java.util.Objects.requireNonNull(siblings.get(metadata.from()), metadata.from());
    }
    if (metadata.factory() != Void.class || !metadata.method().isEmpty()) {
      Class<?> owner = metadata.factory() == Void.class ? component.getType() : metadata.factory();
      try {
        if (metadata.method().isEmpty()) {
          return owner.getConstructor().newInstance();
        }
        var method = owner.getMethod(metadata.method());
        if (!Modifier.isStatic(method.getModifiers())) {
          throw new IllegalStateException("Wire default factories must be static: " + method);
        }
        return java.util.Objects.requireNonNull(method.invoke(null), "default factory result");
      } catch (ReflectiveOperationException failure) {
        throw new IllegalStateException("Cannot resolve wire default for " + component, failure);
      }
    }
    Class<?> type = component.getType();
    String literal = metadata.value();
    if (type == String.class) {
      return literal;
    }
    if (type == boolean.class || type == Boolean.class) {
      if (!"true".equals(literal) && !"false".equals(literal)) {
        throw new IllegalStateException("Boolean default must be true or false: " + component);
      }
      return Boolean.valueOf(literal);
    }
    if (type.isEnum()) {
      return Enum.valueOf((Class<? extends Enum>) type, literal);
    }
    if (type == List.class && (literal.isEmpty() || "[]".equals(literal))) {
      return List.of();
    }
    throw new IllegalStateException("Unsupported literal default for " + component);
  }

  /** Returns a constant default as immutable primitive/map/list metadata when it is not derived. */
  public static Optional<Object> publishedValue(RecordComponent component) {
    WireDefault metadata = component.getAnnotation(WireDefault.class);
    return metadata == null || !metadata.from().isEmpty()
        ? Optional.empty()
        : Optional.of(wireValue(value(component, Map.of())));
  }

  /** Returns a sibling component's canonical position. */
  public static int componentIndex(Class<?> recordType, String name) {
    RecordComponent[] components = recordType.getRecordComponents();
    for (int index = 0; index < components.length; index++) {
      if (components[index].getName().equals(name)) {
        return index;
      }
    }
    throw new IllegalStateException("Unknown default component " + name + " in " + recordType);
  }

  /** Converts a declared default to an immutable JSON-compatible value tree. */
  @SuppressWarnings(
      "PMD.UseConcurrentHashMap") // Thread-confined order is the published record order.
  public static Object wireValue(Object value) {
    if (value instanceof WireScalar text) {
      return text.scalarValue();
    }
    if (value instanceof Enum<?> enumeration) {
      return enumeration.name();
    }
    if (value instanceof String || value instanceof Number || value instanceof Boolean) {
      return value;
    }
    if (value instanceof Map<?, ?> object) {
      Map<String, Object> copy = new LinkedHashMap<>();
      object.forEach((key, field) -> copy.put((String) key, wireValue(field)));
      return Collections.unmodifiableMap(copy);
    }
    if (value instanceof List<?> list) {
      return list.stream().map(ModelDefaults::wireValue).toList();
    }
    if (value.getClass().isRecord()) {
      Map<String, Object> fields = new LinkedHashMap<>();
      addDiscriminator(value.getClass(), fields);
      try {
        for (RecordComponent component : value.getClass().getRecordComponents()) {
          Object field = component.getAccessor().invoke(value);
          if (field instanceof Optional<?> optional) {
            field = optional.orElse(null);
          }
          if (field != null) {
            fields.put(component.getName(), wireValue(field));
          }
        }
      } catch (ReflectiveOperationException failure) {
        throw new IllegalStateException("Cannot publish default " + value.getClass(), failure);
      }
      return Collections.unmodifiableMap(fields);
    }
    throw new IllegalStateException("Unsupported published default " + value.getClass());
  }

  private static void addDiscriminator(Class<?> type, Map<String, Object> fields) {
    Variant variant = type.getAnnotation(Variant.class);
    String property = discriminator(type);
    if (variant != null && !property.isEmpty()) {
      fields.put(property, variant.id());
    }
  }

  private static String discriminator(Class<?> type) {
    Union union = type.getAnnotation(Union.class);
    if (union != null) {
      return union.discriminator();
    }
    for (Class<?> parent : type.getInterfaces()) {
      String discriminator = discriminator(parent);
      if (!discriminator.isEmpty()) {
        return discriminator;
      }
    }
    return "";
  }
}
