package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Validates the static relation between source and persistence intent. */
final class WorkbookStaticPersistenceValidation {
  private WorkbookStaticPersistenceValidation() {}

  static List<WorkbookStaticViolation> validate(WorkbookStaticRequest request) {
    if (request.source().isEmpty() || request.persistence().isEmpty()) {
      return List.of();
    }
    if (request.persistence().orElseThrow() instanceof WorkbookPlan.WorkbookPersistence.Overwrite
        && request.source().orElseThrow() instanceof WorkbookPlan.WorkbookSource.New) {
      return List.of(
          new WorkbookStaticViolation(
              "persistence.type",
              "OVERWRITE persistence requires an EXISTING source; a NEW workbook has no source"
                  + " file to overwrite"));
    }
    return List.of();
  }
}
