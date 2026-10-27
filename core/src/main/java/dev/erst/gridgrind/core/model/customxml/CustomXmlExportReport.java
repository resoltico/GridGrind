package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** XML document exported from one workbook custom-XML mapping. */
@CatalogType(
    id = "CustomXmlExportReport",
    summary =
        "One exported custom-XML mapping payload plus the factual mapping metadata used to produce it.",
    displayOrder = 107)
public record CustomXmlExportReport(
    CustomXmlMappingReport mapping, String encoding, boolean schemaValidated, String xml) {
  public CustomXmlExportReport {
    Objects.requireNonNull(mapping, "mapping must not be null");
    encoding = TextValidation.requireNonBlank(encoding, "encoding");
    xml = TextValidation.requireNonBlank(xml, "xml");
  }
}
