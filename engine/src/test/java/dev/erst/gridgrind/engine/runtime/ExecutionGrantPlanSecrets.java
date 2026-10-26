package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.SecretReference;
import java.util.ArrayList;
import java.util.List;

/** Collects the secret references required by workbook source and persistence security settings. */
final class ExecutionGrantPlanSecrets {
  private ExecutionGrantPlanSecrets() {}

  static List<SecretReference> requestedBy(WorkbookPlan plan) {
    List<SecretReference> references = new ArrayList<>();
    if (plan.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile) {
      existingFile
          .security()
          .flatMap(dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec::passwordRef)
          .ifPresent(references::add);
    }
    persistenceSecurity(plan.persistence()).ifPresent(security -> add(security, references));
    return references.stream().distinct().toList();
  }

  private static java.util.Optional<
          dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec>
      persistenceSecurity(WorkbookPlan.WorkbookPersistence persistence) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> java.util.Optional.empty();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> saveAs.security();
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> overwrite.security();
    };
  }

  private static void add(
      dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec security,
      List<SecretReference> references) {
    if (security.encryption()
        instanceof
        dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.Encrypt encrypt) {
      references.add(encrypt.encryption().passwordRef());
    }
    if (security.signature()
        instanceof dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.Sign sign) {
      references.add(sign.signature().keystorePasswordRef());
      sign.signature().keyPasswordRef().ifPresent(references::add);
    }
  }
}
