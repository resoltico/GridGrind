package dev.erst.gridgrind.excel;

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
    boolean showImportExportValidationErrors,
    boolean autoFit,
    boolean append,
    boolean preserveSortAfLayout,
    boolean preserveFormat,
    Optional<String> schemaNamespace,
    Optional<String> schemaLanguage,
    Optional<String> schemaReference,
    Optional<String> schemaXml,
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
    schemaNamespace = requireNonBlankOptional(schemaNamespace, "schemaNamespace");
    schemaLanguage = requireNonBlankOptional(schemaLanguage, "schemaLanguage");
    schemaReference = requireNonBlankOptional(schemaReference, "schemaReference");
    schemaXml = requireNonBlankOptional(schemaXml, "schemaXml");
    Objects.requireNonNull(dataBinding, "dataBinding must not be null");
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

  private static Optional<String> requireNonBlankOptional(
      Optional<String> value, String fieldName) {
    Optional<String> required = Objects.requireNonNull(value, fieldName + " must not be null");
    required.ifPresent(nonBlank -> requireNonBlank(nonBlank, fieldName));
    return required;
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
