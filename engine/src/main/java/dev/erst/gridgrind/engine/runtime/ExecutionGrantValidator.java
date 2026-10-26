package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.semantics.OperationEffectFootprint;
import dev.erst.gridgrind.core.semantics.OperationSemantics;
import dev.erst.gridgrind.core.semantics.WorkbookOperationContracts;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Evaluates whether one resolved plan fits the authority supplied by its trusted host. */
final class ExecutionGrantValidator {
  private ExecutionGrantValidator() {}

  static List<ExecutionAuthorityDeniedException> violations(
      WorkbookPlan request, GridGrindExecutionGrant grant, Path workingDirectory) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(grant, "grant must not be null");
    Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    GridGrindExecutionGrant.Bounded bounded = (GridGrindExecutionGrant.Bounded) grant;
    List<ExecutionAuthorityDeniedException> violations = new ArrayList<>();
    requireSourceAuthority(request, bounded, workingDirectory, violations);
    requireFormulaEnvironmentAuthority(request, bounded, workingDirectory, violations);
    requireOperationAuthority(request.steps(), bounded, violations);
    requirePublicationAuthority(request, bounded, workingDirectory, violations);
    requireSecretAuthority(request, bounded, violations);
    return List.copyOf(violations);
  }

  private static void requireSecretAuthority(
      WorkbookPlan request,
      GridGrindExecutionGrant.Bounded grant,
      List<ExecutionAuthorityDeniedException> violations) {
    for (SecretReference reference : requestedSecretReferences(request)) {
      if (!grant.allowedSecretReferences().contains(reference)) {
        violations.add(
            new ExecutionAuthorityDeniedException(
                "host grant does not permit secret reference " + reference.id()));
      }
    }
  }

  private static List<SecretReference> requestedSecretReferences(WorkbookPlan request) {
    List<SecretReference> references = new ArrayList<>();
    if (request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile) {
      existingFile
          .security()
          .flatMap(dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec::passwordRef)
          .ifPresent(references::add);
    }
    switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> {}
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          addPersistenceSecrets(saveAs.security(), references);
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite ->
          addPersistenceSecrets(overwrite.security(), references);
    }
    for (WorkbookStep step : request.steps()) {
      if (step instanceof MutationStep mutation) {
        switch (mutation.action()) {
          case dev.erst.gridgrind.core.plan.WorkbookMutationAction.SetSheetProtection protection ->
              protection.passwordRef().ifPresent(references::add);
          case dev.erst.gridgrind.core.plan.WorkbookMutationAction.SetWorkbookProtection
                  protection -> {
            protection.protection().workbookPasswordRef().ifPresent(references::add);
            protection.protection().revisionsPasswordRef().ifPresent(references::add);
          }
          default -> {}
        }
      }
    }
    return references.stream().distinct().toList();
  }

  private static void addPersistenceSecrets(
      java.util.Optional<dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec>
          security,
      List<SecretReference> references) {
    security.ifPresent(
        value -> {
          switch (value.encryption()) {
            case dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.None _ -> {}
            case dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec
                        .PreserveSource
                    _ -> {}
            case dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.Encrypt
                    encrypt ->
                references.add(encrypt.encryption().passwordRef());
          }
          switch (value.signature()) {
            case dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.None _ -> {}
            case dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.Sign sign -> {
              references.add(sign.signature().keystorePasswordRef());
              sign.signature().keyPasswordRef().ifPresent(references::add);
            }
          }
        });
  }

  private static void requireSourceAuthority(
      WorkbookPlan request,
      GridGrindExecutionGrant.Bounded grant,
      Path workingDirectory,
      List<ExecutionAuthorityDeniedException> violations) {
    if (request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile) {
      requireReadableFile(
          grant,
          ExecutionRequestPaths.normalizePath(existingFile.path(), workingDirectory),
          "source workbook",
          violations);
    }
  }

  private static void requireFormulaEnvironmentAuthority(
      WorkbookPlan request,
      GridGrindExecutionGrant.Bounded grant,
      Path workingDirectory,
      List<ExecutionAuthorityDeniedException> violations) {
    for (var externalWorkbook : request.formulaEnvironment().externalWorkbooks()) {
      requireReadableFile(
          grant,
          ExecutionRequestPaths.normalizePath(externalWorkbook.path(), workingDirectory),
          "formula external workbook",
          violations);
    }
  }

  private static void requireReadableFile(
      GridGrindExecutionGrant.Bounded grant,
      Path path,
      String resourceKind,
      List<ExecutionAuthorityDeniedException> violations) {
    boolean allowed =
        grant.readableResources().stream()
            .filter(GridGrindExecutionGrant.ReadAuthority.File.class::isInstance)
            .map(GridGrindExecutionGrant.ReadAuthority.File.class::cast)
            .anyMatch(resource -> resource.path().equals(path));
    if (!allowed) {
      violations.add(
          new ExecutionAuthorityDeniedException(
              "host grant does not permit reading " + resourceKind + ": " + path));
    }
  }

  private static void requireOperationAuthority(
      List<WorkbookStep> steps,
      GridGrindExecutionGrant.Bounded grant,
      List<ExecutionAuthorityDeniedException> violations) {
    for (WorkbookStep step : steps) {
      OperationFact operation = operationFact(step);
      if (!grant.operationIds().contains(operation.id())) {
        violations.add(
            new ExecutionAuthorityDeniedException(
                "host grant does not permit operation " + operation.id()));
      }
      requireTargetAuthority(step, operation, grant, violations);
    }
  }

  private static void requireTargetAuthority(
      WorkbookStep step,
      OperationFact operation,
      GridGrindExecutionGrant.Bounded grant,
      List<ExecutionAuthorityDeniedException> violations) {
    if (!(grant.workbookTargetAuthority()
        instanceof GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets selectedSheets)) {
      return;
    }
    if (operation.semantics().footprint() == OperationEffectFootprint.CONSERVATIVE_WORKBOOK_WIDE) {
      violations.add(
          new ExecutionAuthorityDeniedException(
              "host grant must provide workbook-wide authority for " + operation.id()));
      return;
    }
    var exactSheets = ExecutionTargetScope.exactSheetNames(step.target());
    if (exactSheets.isEmpty()) {
      violations.add(
          new ExecutionAuthorityDeniedException(
              "host grant requires an exactly resolved target footprint for " + operation.id()));
      return;
    }
    for (String sheetName : exactSheets.orElseThrow()) {
      if (!selectedSheets.sheetNames().contains(sheetName)) {
        violations.add(
            new ExecutionAuthorityDeniedException(
                "host grant does not permit target sheet " + sheetName + " for " + operation.id()));
      }
    }
  }

  private static OperationFact operationFact(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutation ->
          new OperationFact(
              mutation.action().actionType(),
              WorkbookOperationContracts.semanticsFor(mutation.action()));
      case AssertionStep assertion ->
          new OperationFact(
              assertion.assertion().assertionType(),
              WorkbookOperationContracts.semanticsFor(assertion.assertion()));
      case InspectionStep inspection ->
          new OperationFact(
              inspection.query().queryType(),
              WorkbookOperationContracts.semanticsFor(inspection.query()));
    };
  }

  private static void requirePublicationAuthority(
      WorkbookPlan request,
      GridGrindExecutionGrant.Bounded grant,
      Path workingDirectory,
      List<ExecutionAuthorityDeniedException> violations) {
    switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> {}
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path requestedPath = ExecutionRequestPaths.normalizePath(saveAs.path(), workingDirectory);
        if (!(grant.publicationAuthority()
                instanceof GridGrindExecutionGrant.PublicationAuthority.SaveAs allowed)
            || !allowed.path().equals(requestedPath)
            || allowed.ifExists() != saveAs.ifExists()) {
          violations.add(
              new ExecutionAuthorityDeniedException(
                  "host grant does not permit SAVE_AS publication to " + requestedPath));
        }
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> {
        if (!(grant.publicationAuthority()
            instanceof GridGrindExecutionGrant.PublicationAuthority.OverwriteSource)) {
          violations.add(
              new ExecutionAuthorityDeniedException(
                  "host grant does not permit overwriting the source workbook"));
        }
      }
    }
  }

  private record OperationFact(String id, OperationSemantics semantics) {
    private OperationFact {
      Objects.requireNonNull(id, "id must not be null");
      Objects.requireNonNull(semantics, "semantics must not be null");
    }
  }
}
