package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

record CatalogTypeDescriptor(
    Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
  CatalogTypeDescriptor {
    Objects.requireNonNull(recordType, "recordType must not be null");
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    optionalFields = CatalogRecordValidation.copyStrings(optionalFields, "optionalFields");
  }

  TypeEntry typeEntry() {
    return CatalogTypeEntryFactory.typeEntry(recordType, id, summary, optionalFields);
  }
}
