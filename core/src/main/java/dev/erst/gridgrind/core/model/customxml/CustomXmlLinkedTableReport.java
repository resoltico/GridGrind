package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Factual table mapping entry linked to one custom-XML map. */
@CatalogType(
    id = "CustomXmlLinkedTableReport",
    summary = "One XML-mapped table linked to a custom-XML mapping.",
    displayOrder = 106)
public record CustomXmlLinkedTableReport(
    SheetLabel sheetName,
    String tableName,
    String tableDisplayName,
    RangeLabel range,
    String commonXPath) {
  public CustomXmlLinkedTableReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    tableName = TextValidation.requireNonBlank(tableName, "tableName");
    tableDisplayName = TextValidation.requireNonBlank(tableDisplayName, "tableDisplayName");
    java.util.Objects.requireNonNull(range, "range must not be null");
    commonXPath = TextValidation.requireNonBlank(commonXPath, "commonXPath");
  }
}
