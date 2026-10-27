package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** Compact summary of one protocol-catalog group for first-contact discovery. */
public record ProtocolCatalogGroupIndex(String group, List<String> entryIds) {
  public ProtocolCatalogGroupIndex {
    group = TextValidation.requireNonBlank(group, "group");
    Objects.requireNonNull(entryIds, "entryIds must not be null");
    entryIds =
        entryIds.stream()
            .map(entryId -> TextValidation.requireNonBlank(entryId, "entryIds entry"))
            .toList();
  }

  /** Number of stable ids published by this group. */
  public int entryCount() {
    return entryIds.size();
  }
}
