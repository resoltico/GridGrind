package dev.erst.gridgrind.contract.catalog;

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
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    noteRefs = CatalogRecordValidation.copyUniqueStrings(noteRefs, "noteRefs");
    projectedFields =
        CatalogRecordValidation.copyProjectedFields(projectedFields, "projectedFields");
  }

  TypeEntry typeEntry() {
    return CatalogTypeEntryFactory.typeEntry(recordType, id, summary, noteRefs, projectedFields);
  }
}
