package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput;
import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises trusted-host admission decisions before workbook or resource access begins. */
class ExecutionGrantValidatorTest {
  @TempDir Path root;

  @Test
  void acceptsAnExactlyAuthorizedSaveAsPlan() {
    WorkbookPlan request = saveAsPlan();
    Path output = root.resolve("report.xlsx");
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("ENSURE_SHEET"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                output, WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    assertTrue(ExecutionGrantValidator.violations(request, grant, root).isEmpty());
  }

  @Test
  void rejectsAnOperationAbsentFromTheHostGrant() {
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                root.resolve("report.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    List<ExecutionAuthorityDeniedException> violations =
        ExecutionGrantValidator.violations(saveAsPlan(), grant, root);

    assertEquals(1, violations.size());
    assertEquals(
        "host grant does not permit operation ENSURE_SHEET", violations.getFirst().getMessage());
  }

  @Test
  void rejectsConservativeWorkbookWideMutationUnderASheetScopedGrant() {
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("ENSURE_SHEET"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets(List.of("Budget")),
            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                root.resolve("report.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    List<ExecutionAuthorityDeniedException> violations =
        ExecutionGrantValidator.violations(saveAsPlan(), grant, root);

    assertEquals(1, violations.size());
    assertEquals(
        "host grant must provide workbook-wide authority for ENSURE_SHEET",
        violations.getFirst().getMessage());
  }

  @Test
  void acceptsTargetBoundReadsWhenTheirSelectorIsContainedInTheGrantedSheets() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new dev.erst.gridgrind.contract.step.InspectionStep(
                    "summary",
                    new SheetSelector.ByName("Budget"),
                    new dev.erst.gridgrind.contract.query.SheetIntrospectionQuery
                        .GetSheetSummary())));
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("GET_SHEET_SUMMARY"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets(List.of("Budget")),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    assertTrue(ExecutionGrantValidator.violations(request, grant, root).isEmpty());
  }

  @Test
  void rejectsTargetBoundReadsWhenTheirSelectorNamesAnUngrantedSheet() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new dev.erst.gridgrind.contract.step.InspectionStep(
                    "summary",
                    new SheetSelector.ByName("Ledger"),
                    new dev.erst.gridgrind.contract.query.SheetIntrospectionQuery
                        .GetSheetSummary())));
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("GET_SHEET_SUMMARY"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets(List.of("Budget")),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    List<ExecutionAuthorityDeniedException> violations =
        ExecutionGrantValidator.violations(request, grant, root);

    assertEquals(1, violations.size());
    assertEquals(
        "host grant does not permit target sheet Ledger for GET_SHEET_SUMMARY",
        violations.getFirst().getMessage());
  }

  @Test
  void rejectsSheetScopedAuthorityWhenTheReadSelectorIsNotExactlyResolved() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new dev.erst.gridgrind.contract.step.InspectionStep(
                    "summary",
                    new SheetSelector.All(),
                    new dev.erst.gridgrind.contract.query.SheetIntrospectionQuery
                        .GetSheetSummary())));
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("GET_SHEET_SUMMARY"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets(List.of("Budget")),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    assertEquals(
        "host grant requires an exactly resolved target footprint for GET_SHEET_SUMMARY",
        ExecutionGrantValidator.violations(request, grant, root).getFirst().getMessage());
  }

  @Test
  void rejectsMismatchedSaveAsAndOverwritePublicationAuthorities() {
    GridGrindExecutionGrant noPublication =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("ENSURE_SHEET"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
    WorkbookPlan overwrite =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(java.util.Optional.empty()),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());

    assertEquals(
        "host grant does not permit SAVE_AS publication to " + root.resolve("report.xlsx"),
        ExecutionGrantValidator.violations(saveAsPlan(), noPublication, root)
            .getLast()
            .getMessage());
    assertEquals(
        "host grant does not permit overwriting the source workbook",
        ExecutionGrantValidator.violations(overwrite, noPublication, root).getLast().getMessage());
    GridGrindExecutionGrant wrongSaveAs =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("ENSURE_SHEET"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                root.resolve("different.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REPLACE),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
    assertEquals(
        "host grant does not permit SAVE_AS publication to " + root.resolve("report.xlsx"),
        ExecutionGrantValidator.violations(saveAsPlan(), wrongSaveAs, root).getLast().getMessage());
    GridGrindExecutionGrant wrongCollisionPolicy =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("ENSURE_SHEET"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                root.resolve("report.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REPLACE),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
    assertEquals(
        "host grant does not permit SAVE_AS publication to " + root.resolve("report.xlsx"),
        ExecutionGrantValidator.violations(saveAsPlan(), wrongCollisionPolicy, root)
            .getLast()
            .getMessage());
  }

  @Test
  void evaluatesEveryPersistenceEncryptionAlternativeForSecretAuthority() {
    WorkbookPlan preservedSourceEncryption =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "report.xlsx",
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                new OoxmlPersistenceSecurityInput(
                    new OoxmlPersistenceEncryptionInput.PreserveSource(),
                    new OoxmlPersistenceSignatureInput.None())),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    GridGrindExecutionGrant noSecretGrant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                root.resolve("report.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    assertTrue(
        ExecutionGrantValidator.violations(preservedSourceEncryption, noSecretGrant, root)
            .isEmpty());
    WorkbookPlan plaintext =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "report.xlsx",
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                OoxmlPersistenceSecurityInput.none()),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    assertTrue(ExecutionGrantValidator.violations(plaintext, noSecretGrant, root).isEmpty());

    SecretReference encryptionPassword = new SecretReference("encryption-password");
    WorkbookPlan encrypted =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "report.xlsx",
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                new OoxmlPersistenceSecurityInput(
                    new OoxmlPersistenceEncryptionInput.Encrypt(
                        new OoxmlEncryptionInput(
                            encryptionPassword,
                            ExcelOoxmlWriteCipher.AES_256,
                            ExcelOoxmlWriteHash.SHA_512)),
                    new OoxmlPersistenceSignatureInput.None())),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    GridGrindExecutionGrant encryptedGrant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                root.resolve("report.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(encryptionPassword),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());

    assertTrue(ExecutionGrantValidator.violations(encrypted, encryptedGrant, root).isEmpty());
  }

  private WorkbookPlan saveAsPlan() {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            "report.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
        List.of(
            new MutationStep(
                "ensure-sheet",
                new SheetSelector.ByName("Budget"),
                new WorkbookMutationAction.EnsureSheet())));
  }
}
