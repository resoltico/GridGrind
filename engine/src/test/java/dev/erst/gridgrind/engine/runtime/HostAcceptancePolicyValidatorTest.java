package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies host preservation fails closed when an authored low-memory mode cannot prove it. */
class HostAcceptancePolicyValidatorTest {
  @Test
  void requiresFullXssfForHostSemanticPreservation() {
    List<HostAcceptancePolicyViolationException> violations =
        HostAcceptancePolicyValidator.violations(
            streamingPlan(),
            new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.PublicationAuthority
                    .None(),
                List.of(),
                preservationPolicy()));

    assertEquals(
        List.of("host semantic preservation requires execution.mode.type=FULL_XSSF"),
        violations.stream().map(Exception::getMessage).toList());
  }

  @Test
  void permitsFullXssfForHostSemanticPreservation() {
    assertTrue(
        HostAcceptancePolicyValidator.violations(
                fullPlan(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded(
                    List.of(),
                    List.of(),
                    new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant
                        .WorkbookTargetAuthority.WorkbookWide(),
                    new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.PublicationAuthority
                        .None(),
                    List.of(),
                    preservationPolicy()))
            .isEmpty());
  }

  @Test
  void requiresFullXssfForHostCalculation() {
    List<HostAcceptancePolicyViolationException> violations =
        HostAcceptancePolicyValidator.violations(
            streamingPlan(),
            new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.PublicationAuthority
                    .None(),
                List.of(),
                new GridGrindHostAcceptancePolicy.RequireAll(
                    List.of(new GridGrindHostAcceptancePolicy.Requirement.RequireCalculation()))));

    assertEquals(
        List.of("host calculation requires execution.mode.type=FULL_XSSF"),
        violations.stream().map(Exception::getMessage).toList());
  }

  @Test
  void failsClosedWhenOpaquePartPreservationCannotProduceAnArtifact() {
    List<HostAcceptancePolicyViolationException> violations =
        HostAcceptancePolicyValidator.violations(
            fullPlan(),
            new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.PublicationAuthority
                    .None(),
                List.of(),
                opaquePartPolicy()));

    assertEquals(
        List.of("host opaque-part preservation requires an unencrypted EXISTING source workbook"),
        violations.stream().map(Exception::getMessage).toList());
  }

  @Test
  void opaquePartPreservationRequiresAnUnencryptedExistingSourceAndPublication() {
    assertEquals(
        List.of("host opaque-part preservation requires execution.mode.type=FULL_XSSF"),
        HostAcceptancePolicyValidator.violations(
                streamingPlan(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded(
                    List.of(),
                    List.of(),
                    new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant
                        .WorkbookTargetAuthority.WorkbookWide(),
                    new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.PublicationAuthority
                        .None(),
                    List.of(),
                    opaquePartPolicy()))
            .stream()
            .map(Exception::getMessage)
            .toList());
    assertEquals(
        List.of("host opaque-part preservation does not support encrypted source workbooks"),
        violationsFor(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.ExistingFile(
                    "input.xlsx",
                    new dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput(
                        java.util.Optional.of(
                            new dev.erst.gridgrind.contract.dto.SecretReference(
                                "source-password")))),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                ExecutionPolicyInput.defaults(),
                FormulaEnvironmentInput.empty(),
                List.of())));
    assertEquals(
        List.of("host opaque-part preservation requires a persisted workbook artifact"),
        violationsFor(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
                new WorkbookPlan.WorkbookPersistence.None(),
                ExecutionPolicyInput.defaults(),
                FormulaEnvironmentInput.empty(),
                List.of())));
    assertTrue(
        violationsFor(
                WorkbookPlan.standard(
                    new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    ExecutionPolicyInput.defaults(),
                    FormulaEnvironmentInput.empty(),
                    List.of()))
            .isEmpty());
  }

  @Test
  void permitsTerminalOnlyAcceptanceAcrossExecutionModes() {
    GridGrindHostAcceptancePolicy policy =
        new GridGrindHostAcceptancePolicy.RequireAll(
            List.of(
                new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(
                    List.of(
                        new AssertionStep(
                            "assert-report",
                            new SheetSelector.ByName("Report"),
                            new PresenceAssertion.SheetPresent())))));

    assertTrue(
        HostAcceptancePolicyValidator.violations(
                streamingPlan(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded(
                    List.of(),
                    List.of(),
                    new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant
                        .WorkbookTargetAuthority.WorkbookWide(),
                    new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.PublicationAuthority
                        .None(),
                    List.of(),
                    policy))
            .isEmpty());
  }

  private static List<String> violationsFor(WorkbookPlan request) {
    return HostAcceptancePolicyValidator.violations(
            request,
            new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.engine.api.GridGrindExecutionGrant.PublicationAuthority
                    .None(),
                List.of(),
                opaquePartPolicy()))
        .stream()
        .map(Exception::getMessage)
        .toList();
  }

  private static WorkbookPlan streamingPlan() {
    return plan(ExecutionPolicyInput.mode(new ExecutionModeInput.StreamingWrite()));
  }

  private static WorkbookPlan fullPlan() {
    return plan(ExecutionPolicyInput.defaults());
  }

  private static WorkbookPlan plan(ExecutionPolicyInput execution) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        execution,
        FormulaEnvironmentInput.empty(),
        List.of());
  }

  private static GridGrindHostAcceptancePolicy preservationPolicy() {
    return new GridGrindHostAcceptancePolicy.RequireAll(
        List.of(
            new GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts(
                List.of(
                    new InspectionStep(
                        "preserve-cell",
                        new CellSelector.ByAddress("Report", "B2"),
                        new SheetIntrospectionQuery.GetCells())))));
  }

  private static GridGrindHostAcceptancePolicy opaquePartPolicy() {
    return new GridGrindHostAcceptancePolicy.RequireAll(
        List.of(
            new GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                List.of("/customXml/item1.xml"))));
  }
}
