package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

record CatalogPlainTypeDescriptor(
    String group,
    Class<? extends Record> recordType,
    String id,
    String summary,
    List<String> noteRefs) {
  CatalogPlainTypeDescriptor(
      String group, Class<? extends Record> recordType, String id, String summary) {
    this(group, recordType, id, summary, List.of());
  }

  CatalogPlainTypeDescriptor {
    group = TextValidation.requireNonBlank(group, "group");
    Objects.requireNonNull(recordType, "recordType must not be null");
    id = TextValidation.requireNonBlank(id, "id");
    summary = TextValidation.requireNonBlank(summary, "summary");
    noteRefs = CatalogRecordValidation.copyUniqueStrings(noteRefs, "noteRefs");
  }

  TypeEntry typeEntry() {
    return CatalogTypeEntryFactory.typeEntry(recordType, id, summary, noteRefs);
  }
}
