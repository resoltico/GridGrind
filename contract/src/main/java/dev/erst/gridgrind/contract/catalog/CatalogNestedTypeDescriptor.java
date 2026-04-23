package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

record CatalogNestedTypeDescriptor(
    String group,
    String discriminatorField,
    Class<?> sealedType,
    List<CatalogTypeDescriptor> typeDescriptors) {
  CatalogNestedTypeDescriptor {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    discriminatorField =
        CatalogRecordValidation.requireNonBlank(discriminatorField, "discriminatorField");
    Objects.requireNonNull(sealedType, "sealedType must not be null");
    typeDescriptors = List.copyOf(typeDescriptors);
    for (CatalogTypeDescriptor typeDescriptor : typeDescriptors) {
      Objects.requireNonNull(typeDescriptor, "typeDescriptors must not contain nulls");
    }
  }
}
