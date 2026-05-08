package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

record CatalogTopLevelTypeDescriptorGroup(
    String group, Class<?> sealedType, List<CatalogTypeDescriptor> typeDescriptors) {
  CatalogTopLevelTypeDescriptorGroup {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    Objects.requireNonNull(sealedType, "sealedType must not be null");
    typeDescriptors = List.copyOf(Objects.requireNonNull(typeDescriptors, "typeDescriptors"));
    if (typeDescriptors.isEmpty()) {
      throw new IllegalArgumentException("typeDescriptors must not be empty");
    }
  }
}
