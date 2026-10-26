package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.json.codec.GridGrindJsonSubtypeSupport;
import java.util.List;
import java.util.Optional;

/** Owns invalid subtype-id wording for the public JSON surface. */
final class GridGrindJsonSubtypeProblemSupport {
  private GridGrindJsonSubtypeProblemSupport() {}

  static Optional<String> specificGuidance(String discriminatorPath, String typeId) {
    if ("source.type".equals(discriminatorPath) && "FILE".equals(typeId)) {
      return Optional.of(
          "use source.type='EXISTING' to open a workbook from disk"
              + " (FILE is only valid for source-backed authored payload inputs)");
    }
    return Optional.empty();
  }

  static Optional<String> specificGuidance(
      String discriminatorPath, String typeId, Optional<Class<?>> similarityRoot) {
    Optional<String> pathSpecific = specificGuidance(discriminatorPath, typeId);
    if (pathSpecific.isPresent() || similarityRoot.isEmpty()) {
      return pathSpecific;
    }
    return Optional.of(
        "valid values: "
            + String.join(", ", GridGrindJsonSubtypeSupport.typeIds(similarityRoot.orElseThrow())));
  }

  static List<String> similarTypeIds(Class<?> baseClass, String typeId) {
    List<String> all = GridGrindJsonSubtypeSupport.typeIds(baseClass);
    String normalized = typeId.toUpperCase(java.util.Locale.ROOT);
    int threshold = Math.min(3, Math.max(1, normalized.length() / 5));
    return all.stream()
        .filter(id -> editDistance(normalized, id) <= threshold)
        .sorted(
            java.util.Comparator.<String>comparingInt(id -> editDistance(normalized, id))
                .thenComparing(java.util.Comparator.naturalOrder()))
        .limit(3)
        .toList();
  }

  @SuppressWarnings("PMD.AvoidArrayLoops")
  private static int editDistance(String a, String b) {
    int m = a.length();
    int n = b.length();
    int[] prev = new int[n + 1];
    int[] curr = new int[n + 1];
    for (int j = 0; j <= n; j++) {
      prev[j] = j;
    }
    for (int i = 1; i <= m; i++) {
      curr[0] = i;
      for (int j = 1; j <= n; j++) {
        if (a.charAt(i - 1) == b.charAt(j - 1)) {
          curr[j] = prev[j - 1];
        } else {
          curr[j] = 1 + Math.min(prev[j - 1], Math.min(prev[j], curr[j - 1]));
        }
      }
      int[] tmp = prev;
      prev = curr;
      curr = tmp;
    }
    return prev[n];
  }
}
