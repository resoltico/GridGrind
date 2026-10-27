package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

record CatalogTypeDescriptor(
    Class<? extends Record> recordType,
    String id,
    String summary,
    List<String> noteRefs,
    List<CatalogProjectedField> projectedFields) {
  CatalogTypeDescriptor(Class<? extends Record> recordType, String id, String summary) {
    this(recordType, id, summary, List.of(), List.of());
  }

  CatalogTypeDescriptor {
    Objects.requireNonNull(recordType, "recordType must not be null");
    id = TextValidation.requireNonBlank(id, "id");
    summary = TextValidation.requireNonBlank(summary, "summary");
    noteRefs = CatalogRecordValidation.copyUniqueStrings(noteRefs, "noteRefs");
    projectedFields =
        CatalogRecordValidation.copyProjectedFields(projectedFields, "projectedFields");
  }

  TypeEntry typeEntry() {
    return CatalogTypeEntryFactory.typeEntry(recordType, id, summary, noteRefs, projectedFields);
  }
}
