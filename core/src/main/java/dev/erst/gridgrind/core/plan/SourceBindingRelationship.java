package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.RecordComponent;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Checks that preparation only binds declared external source leaves in an authored value. */
final class SourceBindingRelationship {
  private SourceBindingRelationship() {}

  static void requireBinding(Object authored, Object bound) {
    if (!matches(authored, bound)) {
      throw new IllegalArgumentException(
          "Prepared inputs may only replace authored external source values");
    }
  }

  private static boolean matches(Object authored, Object bound) {
    if (authored == null || bound == null) return authored == null && bound == null;
    if (authored instanceof TextSourceSpec.Utf8File || authored instanceof TextSourceSpec.Stdin) {
      return bound instanceof TextSourceSpec.Inline;
    }
    if (authored instanceof BinarySourceSpec.File || authored instanceof BinarySourceSpec.Stdin) {
      return bound instanceof BinarySourceSpec.InlineBase64;
    }
    if (authored instanceof List<?> original && bound instanceof List<?> prepared) {
      if (original.size() != prepared.size()) return false;
      for (int index = 0; index < original.size(); index++) {
        if (!matches(original.get(index), prepared.get(index))) return false;
      }
      return true;
    }
    if (authored instanceof Optional<?> original && bound instanceof Optional<?> prepared) {
      return original.isPresent() == prepared.isPresent()
          && (original.isEmpty() || matches(original.orElseThrow(), prepared.orElseThrow()));
    }
    Class<?> type = authored.getClass();
    if (type != bound.getClass()) return false;
    if (!type.isRecord() || !type.getPackageName().startsWith("dev.erst.gridgrind.core."))
      return Objects.equals(authored, bound);
    for (RecordComponent component : type.getRecordComponents()) {
      if (!matches(componentValue(component, authored), componentValue(component, bound))) {
        return false;
      }
    }
    return true;
  }

  private static Object componentValue(RecordComponent component, Object value) {
    try {
      return component.getAccessor().invoke(value);
    } catch (IllegalAccessException | InvocationTargetException failure) {
      throw new IllegalStateException("Cannot inspect a declared core input component", failure);
    }
  }
}
