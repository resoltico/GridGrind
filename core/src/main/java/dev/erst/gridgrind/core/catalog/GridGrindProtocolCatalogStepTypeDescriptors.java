package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import java.util.List;

/** Step-entry descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogStepTypeDescriptors {
  static final List<CatalogTypeDescriptor> STEP_TYPES =
      List.of(
          CatalogTypeEntryFactory.descriptor(
              MutationStep.class,
              "MUTATION",
              "Execute one mutation action against the selected workbook target."),
          CatalogTypeEntryFactory.descriptor(
              AssertionStep.class,
              "ASSERTION",
              "Verify one authored expectation against the selected workbook target."),
          CatalogTypeEntryFactory.descriptor(
              InspectionStep.class,
              "INSPECTION",
              "Run one factual or analytical inspection query against the selected workbook"
                  + " target."));

  private GridGrindProtocolCatalogStepTypeDescriptors() {}
}
