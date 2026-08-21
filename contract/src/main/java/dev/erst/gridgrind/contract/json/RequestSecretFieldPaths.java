package dev.erst.gridgrind.contract.json;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import dev.erst.gridgrind.contract.dto.ProtocolField;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Discovers request-secret ownership paths from the effective creator contract. */
final class RequestSecretFieldPaths {
  private RequestSecretFieldPaths() {}

  static Set<String> forRequestType(Class<?> requestType) {
    Objects.requireNonNull(requestType, "requestType must not be null");
    Set<String> paths = new LinkedHashSet<>();
    collect(requestType, List.of(), paths, new LinkedHashSet<>());
    return Set.copyOf(paths);
  }

  static boolean contains(Set<String> paths, String jsonPath) {
    Objects.requireNonNull(paths, "paths must not be null");
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    return paths.contains(normalize(jsonPath));
  }

  private static void collect(
      Type type, List<String> path, Set<String> paths, Set<Type> activeTypes) {
    if (!activeTypes.add(type)) {
      return;
    }
    try {
      switch (type) {
        case ParameterizedType parameterized ->
            collectParameterized(parameterized, path, paths, activeTypes);
        case Class<?> classType -> collectClass(classType, path, paths, activeTypes);
        default -> {
          // Type variables and wildcards cannot introduce a concrete request-secret owner.
        }
      }
    } finally {
      activeTypes.remove(type);
    }
  }

  private static void collectParameterized(
      ParameterizedType parameterized,
      List<String> path,
      Set<String> paths,
      Set<Type> activeTypes) {
    // The JDK's ParameterizedType contract exposes a class raw type for record components.
    Class<?> rawClass = (Class<?>) parameterized.getRawType();
    if (rawClass == java.util.Optional.class
        || java.util.Collection.class.isAssignableFrom(rawClass)) {
      Type elementType = parameterized.getActualTypeArguments()[0];
      List<String> elementPath =
          rawClass == java.util.Optional.class ? path : appendArraySegment(path);
      collect(elementType, elementPath, paths, activeTypes);
    }
  }

  private static void collectClass(
      Class<?> classType, List<String> path, Set<String> paths, Set<Type> activeTypes) {
    if (classType.isRecord()) {
      collectRecord(classType, path, paths, activeTypes);
    }
    if (classType.isSealed()) {
      for (Class<?> subtype : classType.getPermittedSubclasses()) {
        collect(subtype, path, paths, activeTypes);
      }
    }
  }

  private static void collectRecord(
      Class<?> recordType, List<String> path, Set<String> paths, Set<Type> activeTypes) {
    for (RecordComponent component : recordType.getRecordComponents()) {
      if (component.getAccessor().isAnnotationPresent(JsonIgnore.class)) {
        continue;
      }
      List<String> componentPath =
          append(path, GridGrindProtocolContractSupport.wireFieldName(component));
      ProtocolField field = component.getAnnotation(ProtocolField.class);
      if (field != null && field.secret()) {
        paths.add(render(componentPath));
      } else {
        collect(component.getGenericType(), componentPath, paths, activeTypes);
      }
    }
  }

  private static List<String> append(List<String> path, String segment) {
    return java.util.stream.Stream.concat(path.stream(), java.util.stream.Stream.of(segment))
        .toList();
  }

  private static List<String> appendArraySegment(List<String> path) {
    List<String> result = new java.util.ArrayList<>(path);
    int last = result.size() - 1;
    result.set(last, result.get(last) + "[]");
    return List.copyOf(result);
  }

  private static String render(List<String> path) {
    return String.join(".", path);
  }

  private static String normalize(String jsonPath) {
    return jsonPath.replaceAll("\\[\\d+]", "[]");
  }
}
