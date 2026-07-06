package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

record CatalogPlainTypeDescriptor(
    String group,
    Class<? extends Record> recordType,
    String id,
    String summary,
    List<String> optionalFields,
    List<String> noteRefs) {
  CatalogPlainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    this(group, recordType, id, summary, optionalFields, List.of());
  }

  CatalogPlainTypeDescriptor {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    Objects.requireNonNull(recordType, "recordType must not be null");
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    optionalFields = CatalogRecordValidation.copyStrings(optionalFields, "optionalFields");
    noteRefs = CatalogRecordValidation.copyUniqueStrings(noteRefs, "noteRefs");
  }

  TypeEntry typeEntry() {
    return CatalogTypeEntryFactory.typeEntry(recordType, id, summary, optionalFields, noteRefs);
  }
}
