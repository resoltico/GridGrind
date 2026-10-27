package dev.erst.gridgrind.core.metadata;

import java.lang.reflect.GenericArrayType;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.lang.reflect.WildcardType;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/** Traverses record components and sealed alternatives without scanning a classpath. */
public final class ModelGraph {
  private ModelGraph() {}

  /** Returns each structurally reachable type once, in first-discovery order. */
  public static List<Class<?>> reachableTypes(Type... roots) {
    Set<Class<?>> types = new LinkedHashSet<>();
    for (Type root : roots) {
      collect(root, types);
    }
    return List.copyOf(types);
  }

  private static void collect(Type type, Set<Class<?>> types) {
    if (type instanceof ParameterizedType parameterized) {
      collect(parameterized.getRawType(), types);
      for (Type argument : parameterized.getActualTypeArguments()) {
        collect(argument, types);
      }
    } else if (type instanceof GenericArrayType array) {
      collect(array.getGenericComponentType(), types);
    } else if (type instanceof WildcardType wildcard) {
      for (Type bound : wildcard.getUpperBounds()) {
        collect(bound, types);
      }
    } else if (type instanceof Class<?> classType && types.add(classType)) {
      if (classType.isSealed()) {
        for (Class<?> permitted : classType.getPermittedSubclasses()) {
          collect(permitted, types);
        }
      }
      if (classType.isRecord()) {
        for (var component : classType.getRecordComponents()) {
          collect(component.getGenericType(), types);
        }
      }
    }
  }
}
