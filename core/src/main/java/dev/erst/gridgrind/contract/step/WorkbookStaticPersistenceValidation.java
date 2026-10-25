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
    WorkbookPlan.WorkbookSource source = request.source().orElseThrow();
    WorkbookPlan.WorkbookPersistence persistence = request.persistence().orElseThrow();
    if (persistence instanceof WorkbookPlan.WorkbookPersistence.Overwrite
        && source instanceof WorkbookPlan.WorkbookSource.New) {
      return List.of(
          new WorkbookStaticViolation(
              "persistence.type",
              "OVERWRITE persistence requires an EXISTING source; a NEW workbook has no source"
                  + " file to overwrite"));
    }
    java.util.Optional<dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput> security =
        persistenceSecurity(persistence);
    if (source instanceof WorkbookPlan.WorkbookSource.ExistingFile
        && isWriting(persistence)
        && security.isEmpty()) {
      return List.of(
          new WorkbookStaticViolation(
              "persistence.security",
              "EXISTING-source writes must declare persistence.security.encryption and"
                  + " persistence.security.signature explicitly"));
    }
    return List.of();
  }

  private static boolean isWriting(WorkbookPlan.WorkbookPersistence persistence) {
    return !(persistence instanceof WorkbookPlan.WorkbookPersistence.None);
  }

  private static java.util.Optional<dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput>
      persistenceSecurity(WorkbookPlan.WorkbookPersistence persistence) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> java.util.Optional.empty();
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> overwrite.security();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> saveAs.security();
    };
  }
}
