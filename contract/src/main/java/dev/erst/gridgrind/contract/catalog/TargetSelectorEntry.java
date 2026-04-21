package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/** Machine-readable description of one allowed target selector family for an operation entry. */
public record TargetSelectorEntry(String family, List<String> typeIds) {
  public TargetSelectorEntry {
    family = CatalogRecordValidation.requireNonBlank(family, "family");
    typeIds = CatalogRecordValidation.copyStrings(typeIds, "typeIds");
  }
}
