package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.List;

/** Owns the concrete protocol entry descriptors published by the public catalog. */
final class GridGrindProtocolCatalogTypeDescriptors {
  static final List<CatalogTypeDescriptor> STEP_TYPES =
      GridGrindProtocolCatalogStepTypeDescriptors.STEP_TYPES;

  static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      GridGrindProtocolCatalogSourceTypeDescriptors.SOURCE_TYPES;

  static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      GridGrindProtocolCatalogPersistenceTypeDescriptors.PERSISTENCE_TYPES;

  static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      ProtocolTypeMetadataSupport.catalogDescriptorsFor(MutationAction.class);

  static final List<CatalogTypeDescriptor> ASSERTION_TYPES =
      ProtocolTypeMetadataSupport.catalogDescriptorsFor(Assertion.class);

  static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      ProtocolTypeMetadataSupport.catalogDescriptorsFor(InspectionQuery.class);

  static final List<CatalogTopLevelTypeDescriptorGroup> TOP_LEVEL_GROUPS =
      List.of(
          new CatalogTopLevelTypeDescriptorGroup(
              "sourceTypes", WorkbookPlan.WorkbookSource.class, SOURCE_TYPES),
          new CatalogTopLevelTypeDescriptorGroup(
              "persistenceTypes", WorkbookPlan.WorkbookPersistence.class, PERSISTENCE_TYPES),
          new CatalogTopLevelTypeDescriptorGroup("stepTypes", WorkbookStep.class, STEP_TYPES),
          new CatalogTopLevelTypeDescriptorGroup(
              "mutationActionTypes", MutationAction.class, MUTATION_ACTION_TYPES),
          new CatalogTopLevelTypeDescriptorGroup(
              "assertionTypes", Assertion.class, ASSERTION_TYPES),
          new CatalogTopLevelTypeDescriptorGroup(
              "inspectionQueryTypes", InspectionQuery.class, INSPECTION_QUERY_TYPES));

  static final List<CatalogTypeDescriptor> ALL_TYPES =
      TOP_LEVEL_GROUPS.stream()
          .map(CatalogTopLevelTypeDescriptorGroup::typeDescriptors)
          .flatMap(List::stream)
          .toList();

  private GridGrindProtocolCatalogTypeDescriptors() {}
}
