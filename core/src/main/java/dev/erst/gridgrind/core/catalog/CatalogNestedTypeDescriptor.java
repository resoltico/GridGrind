package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

record CatalogNestedTypeDescriptor(
    String group,
    String discriminatorField,
    Class<?> sealedType,
    List<CatalogTypeDescriptor> typeDescriptors) {
  CatalogNestedTypeDescriptor {
    group = TextValidation.requireNonBlank(group, "group");
    discriminatorField = TextValidation.requireNonBlank(discriminatorField, "discriminatorField");
    Objects.requireNonNull(sealedType, "sealedType must not be null");
    typeDescriptors = List.copyOf(typeDescriptors);
    for (CatalogTypeDescriptor typeDescriptor : typeDescriptors) {
      Objects.requireNonNull(typeDescriptor, "typeDescriptors must not contain nulls");
    }
  }
}
