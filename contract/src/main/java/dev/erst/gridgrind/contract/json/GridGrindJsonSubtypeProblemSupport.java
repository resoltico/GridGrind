package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Optional;

/** Owns invalid subtype-id wording for the public JSON surface. */
final class GridGrindJsonSubtypeProblemSupport {
  private GridGrindJsonSubtypeProblemSupport() {}

  static String unknownTypeValueMessage(
      tools.jackson.databind.exc.InvalidTypeIdException exception) {
    String typeId = exception.getTypeId();
    if (typeId == null) {
      return GridGrindJsonValueProblemSupport.productOwnedJacksonMessage(
          GridGrindJsonProblemMessageSupport.cleanJacksonMessage(exception.getOriginalMessage()));
    }
    String defaultMessage = "Unknown type value '" + typeId + "'";
    Optional<String> containerName =
        GridGrindJsonPayloadMetadataSupport.terminalContainerName(exception.getPath());
    if (containerName.isPresent() && "source".equals(containerName.orElseThrow())) {
      if ("FILE".equals(typeId)) {
        return defaultMessage
            + "; use source.type='EXISTING' to open a workbook from disk"
            + " (FILE is only valid for source-backed authored payload inputs)";
      }
      return withCandidates(defaultMessage, exception, typeId);
    }
    return withCandidates(defaultMessage, exception, typeId);
  }

  private static String withCandidates(
      String base, tools.jackson.databind.exc.InvalidTypeIdException exception, String typeId) {
    tools.jackson.databind.JavaType baseType = exception.getBaseType();
    if (baseType == null) {
      return base;
    }
    List<String> candidates = similarTypeIds(baseType.getRawClass(), typeId);
    if (candidates.isEmpty()) {
      return base;
    }
    return base + "; similar valid values: " + String.join(", ", candidates);
  }

  private static List<String> similarTypeIds(Class<?> baseClass, String typeId) {
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
