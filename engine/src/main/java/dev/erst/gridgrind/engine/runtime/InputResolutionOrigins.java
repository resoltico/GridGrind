package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.WireScalar;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.RecordComponent;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Maps each bound source-input object to its authored request field without value-based guessing.
 */
@SuppressWarnings(
    "PMD.LooseCoupling") // Identity keys distinguish equal source records at distinct paths.
final class InputResolutionOrigins {
  private final IdentityHashMap<Object, java.util.ArrayDeque<JsonLocation>> locations;
  private final IdentityHashMap<Object, JsonLocation> recent = new IdentityHashMap<>();

  private InputResolutionOrigins(
      IdentityHashMap<Object, java.util.ArrayDeque<JsonLocation>> locations) {
    this.locations = new IdentityHashMap<>(locations);
  }

  static InputResolutionOrigins forRequest(
      WorkbookPlan request, Optional<RequestIntake> requestAnalysis) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(requestAnalysis, "requestAnalysis must not be null");
    IdentityHashMap<Object, java.util.ArrayDeque<JsonLocation>> locations = new IdentityHashMap<>();
    collect(request, "", requestAnalysis, locations);
    return new InputResolutionOrigins(locations);
  }

  static InputResolutionOrigins forValue(Object value, String path) {
    IdentityHashMap<Object, java.util.ArrayDeque<JsonLocation>> locations = new IdentityHashMap<>();
    collect(value, path, Optional.empty(), locations);
    return new InputResolutionOrigins(locations);
  }

  InputResolutionOrigins scoped(String path) {
    IdentityHashMap<Object, java.util.ArrayDeque<JsonLocation>> selected = new IdentityHashMap<>();
    locations.forEach((source, occurrences) -> {
      var matching = new java.util.ArrayDeque<JsonLocation>();
      for (var occurrence : occurrences) {
        if (occurrence
            .jsonPathValue()
            .filter(candidate ->
                candidate.equals(path)
                    || candidate.startsWith(path + ".")
                    || candidate.startsWith(path + "["))
            .isPresent()) {
          matching.add(occurrence);
        }
      }
      if (!matching.isEmpty()) selected.put(source, matching);
    });
    return new InputResolutionOrigins(selected);
  }

  Optional<JsonLocation> locationFor(Object source) {
    var occurrences = locations.get(Objects.requireNonNull(source, "source must not be null"));
    JsonLocation location = occurrences == null ? null : occurrences.pollFirst();
    if (location != null) recent.put(source, location);
    return Optional.ofNullable(location);
  }

  Optional<JsonLocation> lastLocationFor(Object source) {
    return Optional.ofNullable(recent.get(source));
  }

  private static void collect(
      Object value,
      String jsonPath,
      Optional<RequestIntake> requestAnalysis,
      IdentityHashMap<Object, java.util.ArrayDeque<JsonLocation>> locations) {
    if (value == null || value instanceof WireScalar) {
      return;
    }
    if (value instanceof TextSourceSpec textSource) {
      locations
          .computeIfAbsent(textSource, ignored -> new java.util.ArrayDeque<>())
          .add(locationFor(textSource, jsonPath, requestAnalysis));
      return;
    }
    if (value instanceof BinarySourceSpec binarySource) {
      locations
          .computeIfAbsent(binarySource, ignored -> new java.util.ArrayDeque<>())
          .add(locationFor(binarySource, jsonPath, requestAnalysis));
      return;
    }
    if (value instanceof Optional<?> optional) {
      optional.ifPresent(candidate -> collect(candidate, jsonPath, requestAnalysis, locations));
      return;
    }
    if (value instanceof List<?> values) {
      for (int index = 0; index < values.size(); index++) {
        collect(values.get(index), jsonPath + "[" + index + "]", requestAnalysis, locations);
      }
      return;
    }
    if (!value.getClass().isRecord()) {
      return;
    }
    for (RecordComponent component : value.getClass().getRecordComponents()) {
      collect(
          componentValue(value, component),
          append(jsonPath, jsonFieldName(component)),
          requestAnalysis,
          locations);
    }
  }

  private static JsonLocation locationFor(
      TextSourceSpec source, String sourcePath, Optional<RequestIntake> requestAnalysis) {
    return locationFor(sourcePath, sourceField(source), requestAnalysis);
  }

  private static JsonLocation locationFor(
      BinarySourceSpec source, String sourcePath, Optional<RequestIntake> requestAnalysis) {
    return locationFor(sourcePath, sourceField(source), requestAnalysis);
  }

  private static JsonLocation locationFor(
      String sourcePath, String fieldName, Optional<RequestIntake> requestAnalysis) {
    String jsonPath = append(sourcePath, fieldName);
    return requestAnalysis
        .map(analysis -> analysis.jsonLocationAt(jsonPath))
        .orElseGet(() -> JsonLocation.pathOnly(jsonPath));
  }

  static String sourceField(TextSourceSpec source) {
    return switch (source) {
      case TextSourceSpec.Inline _ -> "text";
      case TextSourceSpec.Utf8File _ -> "path";
      case TextSourceSpec.Stdin _ -> "type";
    };
  }

  static String sourceField(BinarySourceSpec source) {
    return switch (source) {
      case BinarySourceSpec.InlineBase64 _ -> "base64Data";
      case BinarySourceSpec.File _ -> "path";
      case BinarySourceSpec.Stdin _ -> "type";
    };
  }

  static Object componentValue(Object record, RecordComponent component) {
    try {
      return component.getAccessor().invoke(record);
    } catch (IllegalAccessException | InvocationTargetException exception) {
      throw new IllegalStateException(
          "Unable to inspect request component " + component.getName(), exception);
    }
  }

  static String jsonFieldName(RecordComponent component) {
    return dev.erst.gridgrind.core.catalog.GridGrindProtocolContractSupport.wireFieldName(
        component);
  }

  private static String append(String prefix, String fieldName) {
    return prefix.isEmpty() ? fieldName : prefix + "." + fieldName;
  }
}
