package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;

/** Machine-readable description of one allowed target selector family for an operation entry. */
public record TargetSelectorEntry(String family, List<String> typeIds) {
  public TargetSelectorEntry {
    family = TextValidation.requireNonBlank(family, "family");
    typeIds = CatalogRecordValidation.copyStrings(typeIds, "typeIds");
  }
}
