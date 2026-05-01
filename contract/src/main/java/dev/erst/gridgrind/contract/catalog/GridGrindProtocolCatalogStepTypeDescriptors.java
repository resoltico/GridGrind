package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.List;

/** Step-entry descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogStepTypeDescriptors {
  static final List<CatalogTypeDescriptor> STEP_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              MutationStep.class,
              "MUTATION",
              "Execute one mutation action against the selected workbook target."),
          GridGrindProtocolCatalog.descriptor(
              AssertionStep.class,
              "ASSERTION",
              "Verify one authored expectation against the selected workbook target."),
          GridGrindProtocolCatalog.descriptor(
              InspectionStep.class,
              "INSPECTION",
              "Run one factual or analytical inspection query against the selected workbook"
                  + " target."));

  private GridGrindProtocolCatalogStepTypeDescriptors() {}
}
