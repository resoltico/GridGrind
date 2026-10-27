package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Factual workbook custom-XML data-binding metadata returned by inspection. */
@CatalogType(
    id = "CustomXmlDataBindingReport",
    summary = "Optional custom-XML data-binding metadata attached to one workbook mapping.",
    displayOrder = 104)
public record CustomXmlDataBindingReport(
    Optional<String> dataBindingName,
    Optional<Boolean> fileBinding,
    Optional<Long> connectionId,
    Optional<String> fileBindingName,
    long loadMode) {
  public CustomXmlDataBindingReport {
    dataBindingName =
        Objects.requireNonNull(dataBindingName, "dataBindingName must not be null")
            .map(value -> TextValidation.requireNonBlank(value, "dataBindingName"));
    Objects.requireNonNull(fileBinding, "fileBinding must not be null");
    Objects.requireNonNull(connectionId, "connectionId must not be null");
    if (connectionId.filter(value -> value < 0L).isPresent()) {
      throw new IllegalArgumentException("connectionId must not be negative");
    }
    fileBindingName =
        Objects.requireNonNull(fileBindingName, "fileBindingName must not be null")
            .map(value -> TextValidation.requireNonBlank(value, "fileBindingName"));
    if (loadMode < 0L) {
      throw new IllegalArgumentException("loadMode must not be negative");
    }
  }
}
