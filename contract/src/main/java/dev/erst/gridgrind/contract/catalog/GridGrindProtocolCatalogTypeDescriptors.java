package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.stream.Stream;

/** Owns the concrete protocol entry descriptors published by the public catalog. */
final class GridGrindProtocolCatalogTypeDescriptors {
  static final List<CatalogTypeDescriptor> STEP_TYPES =
      GridGrindProtocolCatalogStepTypeDescriptors.STEP_TYPES;

  static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      GridGrindProtocolCatalogSourceTypeDescriptors.SOURCE_TYPES;

  static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      GridGrindProtocolCatalogPersistenceTypeDescriptors.PERSISTENCE_TYPES;

  static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      GridGrindProtocolCatalogMutationActionTypeDescriptors.MUTATION_ACTION_TYPES;

  static final List<CatalogTypeDescriptor> ASSERTION_TYPES =
      GridGrindProtocolCatalogAssertionTypeDescriptors.ASSERTION_TYPES;

  static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      GridGrindProtocolCatalogInspectionQueryTypeDescriptors.INSPECTION_QUERY_TYPES;

  static final List<CatalogTypeDescriptor> ALL_TYPES =
      Stream.of(
              STEP_TYPES,
              SOURCE_TYPES,
              PERSISTENCE_TYPES,
              MUTATION_ACTION_TYPES,
              ASSERTION_TYPES,
              INSPECTION_QUERY_TYPES)
          .flatMap(List::stream)
          .toList();

  private GridGrindProtocolCatalogTypeDescriptors() {}
}
