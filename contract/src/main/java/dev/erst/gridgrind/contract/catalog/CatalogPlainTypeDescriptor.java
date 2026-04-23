package dev.erst.gridgrind.contract.catalog;

import java.util.Objects;

record CatalogPlainTypeDescriptor(
    String group, Class<? extends Record> recordType, TypeEntry typeEntry) {
  CatalogPlainTypeDescriptor {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    Objects.requireNonNull(recordType, "recordType must not be null");
    Objects.requireNonNull(typeEntry, "typeEntry must not be null");
  }
}
