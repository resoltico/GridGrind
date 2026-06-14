package dev.erst.gridgrind.excel.customxml;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Immutable factual workbook custom-XML mapping metadata. */
public record ExcelCustomXmlMappingSnapshot(
    long mapId,
    String name,
    String rootElement,
    String schemaId,
    ExcelCustomXmlMappingSettings settings,
    ExcelCustomXmlSchemaSnapshot schema,
    Optional<ExcelCustomXmlDataBindingSnapshot> dataBinding,
    List<ExcelCustomXmlLinkedCellSnapshot> linkedCells,
    List<ExcelCustomXmlLinkedTableSnapshot> linkedTables) {
  public ExcelCustomXmlMappingSnapshot {
    if (mapId <= 0L) {
      throw new IllegalArgumentException("mapId must be greater than 0");
    }
    name = requireNonBlank(name, "name");
    rootElement = requireNonBlank(rootElement, "rootElement");
    schemaId = requireNonBlank(schemaId, "schemaId");
    Objects.requireNonNull(settings, "settings must not be null");
    Objects.requireNonNull(schema, "schema must not be null");
    dataBinding = Objects.requireNonNullElseGet(dataBinding, Optional::empty);
    linkedCells = copyValues(linkedCells, "linkedCells");
    linkedTables = copyValues(linkedTables, "linkedTables");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copy);
  }
}
