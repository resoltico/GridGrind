package dev.erst.gridgrind.contract.dto;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual workbook custom-XML mapping metadata returned by inspection. */
public record CustomXmlMappingReport(
    long mapId,
    String name,
    String rootElement,
    String schemaId,
    boolean showImportExportValidationErrors,
    boolean autoFit,
    boolean append,
    boolean preserveSortAfLayout,
    boolean preserveFormat,
    String schemaNamespace,
    String schemaLanguage,
    String schemaReference,
    String schemaXml,
    CustomXmlDataBindingReport dataBinding,
    List<CustomXmlLinkedCellReport> linkedCells,
    List<CustomXmlLinkedTableReport> linkedTables) {
  public CustomXmlMappingReport {
    if (mapId <= 0L) {
      throw new IllegalArgumentException("mapId must be greater than 0");
    }
    name = requireNonBlank(name, "name");
    rootElement = requireNonBlank(rootElement, "rootElement");
    schemaId = requireNonBlank(schemaId, "schemaId");
    if (schemaNamespace != null) {
      schemaNamespace = requireNonBlank(schemaNamespace, "schemaNamespace");
    }
    if (schemaLanguage != null) {
      schemaLanguage = requireNonBlank(schemaLanguage, "schemaLanguage");
    }
    if (schemaReference != null) {
      schemaReference = requireNonBlank(schemaReference, "schemaReference");
    }
    if (schemaXml != null) {
      schemaXml = requireNonBlank(schemaXml, "schemaXml");
    }
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
