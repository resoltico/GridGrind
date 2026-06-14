package dev.erst.gridgrind.contract.dto;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Factual workbook custom-XML mapping metadata returned by inspection. */
public record CustomXmlMappingReport(
    long mapId,
    String name,
    String rootElement,
    String schemaId,
    Settings settings,
    Schema schema,
    @Nullable CustomXmlDataBindingReport dataBinding,
    List<CustomXmlLinkedCellReport> linkedCells,
    List<CustomXmlLinkedTableReport> linkedTables) {
  public CustomXmlMappingReport {
    if (mapId <= 0L) {
      throw new IllegalArgumentException("mapId must be greater than 0");
    }
    name = requireNonBlank(name, "name");
    rootElement = requireNonBlank(rootElement, "rootElement");
    schemaId = requireNonBlank(schemaId, "schemaId");
    Objects.requireNonNull(settings, "settings must not be null");
    Objects.requireNonNull(schema, "schema must not be null");
    linkedCells = copyValues(linkedCells, "linkedCells");
    linkedTables = copyValues(linkedTables, "linkedTables");
  }

  /** Persisted custom-XML map behavior flags. */
  public record Settings(
      boolean showImportExportValidationErrors,
      boolean autoFit,
      boolean append,
      boolean preserveSortAfLayout,
      boolean preserveFormat) {}

  /** Optional schema metadata attached to one custom-XML mapping. */
  public record Schema(
      @Nullable String namespace,
      @Nullable String language,
      @Nullable String reference,
      @Nullable String xml) {
    public Schema {
      if (namespace != null) {
        namespace = requireNonBlank(namespace, "namespace");
      }
      if (language != null) {
        language = requireNonBlank(language, "language");
      }
      if (reference != null) {
        reference = requireNonBlank(reference, "reference");
      }
      if (xml != null) {
        xml = requireNonBlank(xml, "xml");
      }
    }
  }

  private static String requireNonBlank(@Nullable String value, String fieldName) {
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
