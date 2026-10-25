package dev.erst.gridgrind.contract.catalog;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Machine-readable example showing how one operation is placed inside a workbook step. */
public record ProtocolStepTemplate(
    String stepKind, String bodyField, Map<String, Object> template, List<String> notes) {
  public ProtocolStepTemplate {
    stepKind = CatalogRecordValidation.requireNonBlank(stepKind, "stepKind");
    bodyField = CatalogRecordValidation.requireNonBlank(bodyField, "bodyField");
    template = copyObject(Objects.requireNonNull(template, "template must not be null"));
    notes = CatalogRecordValidation.copyStrings(notes, "notes");
  }

  // Preserve authored field order in an immutable copy; no mutable map escapes.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<String, Object> copyObject(Map<?, ?> source) {
    Map<String, Object> copy = new LinkedHashMap<>();
    source.forEach(
        (key, value) -> {
          if (!(key instanceof String field)) {
            throw new IllegalArgumentException("Template fields must have string names");
          }
          copy.put(field, copyValue(value));
        });
    return Collections.unmodifiableMap(copy);
  }

  private static Object copyValue(Object value) {
    return switch (value) {
      case String text -> text;
      case Byte number -> number;
      case Short number -> number;
      case Integer number -> number;
      case Long number -> number;
      case Float number -> number;
      case Double number -> number;
      case java.math.BigInteger number -> number;
      case java.math.BigDecimal number -> number;
      case Boolean flag -> flag;
      case Map<?, ?> object -> copyObject(object);
      case List<?> list -> list.stream().map(ProtocolStepTemplate::copyValue).toList();
      default -> throw new IllegalArgumentException("Unsupported template value: " + value);
    };
  }
}
