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
  private final IdentityHashMap<Object, JsonLocation> locations;

  private InputResolutionOrigins(IdentityHashMap<Object, JsonLocation> locations) {
    this.locations = new IdentityHashMap<>(locations);
  }

  static InputResolutionOrigins forRequest(
      WorkbookPlan request, Optional<RequestIntake> requestAnalysis) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(requestAnalysis, "requestAnalysis must not be null");
    IdentityHashMap<Object, JsonLocation> locations = new IdentityHashMap<>();
    collect(request, "", requestAnalysis, locations);
    return new InputResolutionOrigins(locations);
  }

  Optional<JsonLocation> locationFor(Object source) {
    return Optional.ofNullable(
        locations.get(Objects.requireNonNull(source, "source must not be null")));
  }

  private static void collect(
      Object value,
      String jsonPath,
      Optional<RequestIntake> requestAnalysis,
      IdentityHashMap<Object, JsonLocation> locations) {
    if (value == null || value instanceof WireScalar) {
      return;
    }
    if (value instanceof TextSourceSpec textSource) {
      locations.put(textSource, locationFor(textSource, jsonPath, requestAnalysis));
      return;
    }
    if (value instanceof BinarySourceSpec binarySource) {
      locations.put(binarySource, locationFor(binarySource, jsonPath, requestAnalysis));
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
