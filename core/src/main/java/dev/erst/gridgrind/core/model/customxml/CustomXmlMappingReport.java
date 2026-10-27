package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual workbook custom-XML mapping metadata returned by inspection. */
@CatalogType(
    id = "CustomXmlMappingReport",
    summary = "One factual workbook custom-XML mapping report.",
    displayOrder = 101)
public record CustomXmlMappingReport(
    long mapId,
    String name,
    String rootElement,
    String schemaId,
    Settings settings,
    Schema schema,
    Optional<CustomXmlDataBindingReport> dataBinding,
    List<CustomXmlLinkedCellReport> linkedCells,
    List<CustomXmlLinkedTableReport> linkedTables) {
  public CustomXmlMappingReport {
    if (mapId <= 0L) {
      throw new IllegalArgumentException("mapId must be greater than 0");
    }
    name = TextValidation.requireNonBlank(name, "name");
    rootElement = TextValidation.requireNonBlank(rootElement, "rootElement");
    schemaId = TextValidation.requireNonBlank(schemaId, "schemaId");
    Objects.requireNonNull(settings, "settings must not be null");
    Objects.requireNonNull(schema, "schema must not be null");
    Objects.requireNonNull(dataBinding, "dataBinding must not be null");
    linkedCells = copyValues(linkedCells, "linkedCells");
    linkedTables = copyValues(linkedTables, "linkedTables");
  }

  /** Persisted custom-XML map behavior flags. */
  @CatalogType(
      id = "CustomXmlMappingSettingsReport",
      summary = "Persisted custom-XML map behavior flags.",
      displayOrder = 102)
  public record Settings(
      boolean showImportExportValidationErrors,
      boolean autoFit,
      boolean append,
      boolean preserveSortAfLayout,
      boolean preserveFormat) {}

  /** Optional schema metadata attached to one custom-XML mapping. */
  @CatalogType(
      id = "CustomXmlMappingSchemaReport",
      summary = "Optional schema metadata attached to one custom-XML mapping.",
      displayOrder = 103)
  public record Schema(
      Optional<String> namespace,
      Optional<String> language,
      Optional<String> reference,
      Optional<String> xml) {
    public Schema {
      namespace =
          Objects.requireNonNull(namespace, "namespace must not be null")
              .map(value -> TextValidation.requireNonBlank(value, "namespace"));
      language =
          Objects.requireNonNull(language, "language must not be null")
              .map(value -> TextValidation.requireNonBlank(value, "language"));
      reference =
          Objects.requireNonNull(reference, "reference must not be null")
              .map(value -> TextValidation.requireNonBlank(value, "reference"));
      xml =
          Objects.requireNonNull(xml, "xml must not be null")
              .map(value -> TextValidation.requireNonBlank(value, "xml"));
    }
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
