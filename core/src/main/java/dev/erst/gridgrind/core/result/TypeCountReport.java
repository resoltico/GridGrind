package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;

/** Count of one observed cell value type inside a schema column. */
public record TypeCountReport(String type, int count) {
  private static final List<String> SUPPORTED_SCHEMA_TYPES =
      List.of("TEXT", "NUMBER", "BOOLEAN", "ERROR", "DATE", "TIME", "DATE_TIME");

  public TypeCountReport {
    type = TextValidation.requireNonBlank(type, "type");
    requireSupportedType(type, "type");
    if (count <= 0) {
      throw new IllegalArgumentException("count must be greater than 0");
    }
  }

  /** Validates that a schema type name uses the published cell-readback vocabulary. */
  public static void requireSupportedType(String type, String fieldName) {
    if (!SUPPORTED_SCHEMA_TYPES.contains(type)) {
      throw new IllegalArgumentException(
          fieldName
              + " must be one of "
              + String.join(", ", SUPPORTED_SCHEMA_TYPES)
              + " but was "
              + type);
    }
  }
}
