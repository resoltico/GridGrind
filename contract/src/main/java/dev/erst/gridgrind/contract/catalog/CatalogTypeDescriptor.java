package dev.erst.gridgrind.contract.catalog;

import java.util.Objects;

record CatalogTypeDescriptor(Class<? extends Record> recordType, TypeEntry typeEntry) {
  CatalogTypeDescriptor {
    Objects.requireNonNull(recordType, "recordType must not be null");
    Objects.requireNonNull(typeEntry, "typeEntry must not be null");
  }
}
