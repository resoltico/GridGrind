package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies host preservation fails closed when an authored low-memory mode cannot prove it. */
class HostAcceptancePolicyValidatorTest {
  @Test
  void requiresFullXssfForHostSemanticPreservation() {
    List<HostAcceptancePolicyViolationException> violations =
        HostAcceptancePolicyValidator.violations(
            streamingPlan(),
            new dev.erst.gridgrind.core.grant.ExecutionGrant(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
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
                new dev.erst.gridgrind.core.grant.ExecutionGrant(
                    List.of(),
                    List.of(),
                    new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                        .WorkbookWide(),
                    new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
                    List.of(),
                    preservationPolicy()))
            .isEmpty());
  }

  @Test
  void requiresFullXssfForHostCalculation() {
    List<HostAcceptancePolicyViolationException> violations =
        HostAcceptancePolicyValidator.violations(
            streamingPlan(),
            new dev.erst.gridgrind.core.grant.ExecutionGrant(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                new AcceptancePolicy.RequireAll(
                    List.of(new AcceptancePolicy.Requirement.RequireCalculation()))));

    assertEquals(
        List.of("host calculation requires execution.mode.type=FULL_XSSF"),
        violations.stream().map(Exception::getMessage).toList());
  }

  @Test
  void failsClosedWhenOpaquePartPreservationCannotProduceAnArtifact() {
    List<HostAcceptancePolicyViolationException> violations =
        HostAcceptancePolicyValidator.violations(
            fullPlan(),
            new dev.erst.gridgrind.core.grant.ExecutionGrant(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
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
                new dev.erst.gridgrind.core.grant.ExecutionGrant(
                    List.of(),
                    List.of(),
                    new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                        .WorkbookWide(),
                    new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
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
                    new dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec(
                        java.util.Optional.of(
                            new dev.erst.gridgrind.core.source.SecretReference(
                                "source-password")))),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                ExecutionPolicySpec.defaults(),
                FormulaEnvironmentSpec.empty(),
                List.of())));
    assertEquals(
        List.of("host opaque-part preservation requires a persisted workbook artifact"),
        violationsFor(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
                new WorkbookPlan.WorkbookPersistence.None(),
                ExecutionPolicySpec.defaults(),
                FormulaEnvironmentSpec.empty(),
                List.of())));
    assertTrue(
        violationsFor(
                WorkbookPlan.standard(
                    new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    ExecutionPolicySpec.defaults(),
                    FormulaEnvironmentSpec.empty(),
                    List.of()))
            .isEmpty());
  }

  @Test
  void permitsTerminalOnlyAcceptanceAcrossExecutionModes() {
    AcceptancePolicy policy =
        new AcceptancePolicy.RequireAll(
            List.of(
                new AcceptancePolicy.Requirement.TerminalAssertions(
                    List.of(
                        new AssertionStep(
                            new StepId("assert-report"),
                            new SheetSelector.ByName(new SheetName("Report")),
                            new PresenceAssertion.SheetPresent())))));

    assertTrue(
        HostAcceptancePolicyValidator.violations(
                streamingPlan(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant(
                    List.of(),
                    List.of(),
                    new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                        .WorkbookWide(),
                    new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
                    List.of(),
                    policy))
            .isEmpty());
  }

  private static List<String> violationsFor(WorkbookPlan request) {
    return HostAcceptancePolicyValidator.violations(
            request,
            new dev.erst.gridgrind.core.grant.ExecutionGrant(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                opaquePartPolicy()))
        .stream()
        .map(Exception::getMessage)
        .toList();
  }

  private static WorkbookPlan streamingPlan() {
    return plan(ExecutionPolicySpec.mode(new ExecutionModeSpec.StreamingWrite()));
  }

  private static WorkbookPlan fullPlan() {
    return plan(ExecutionPolicySpec.defaults());
  }

  private static WorkbookPlan plan(ExecutionPolicySpec execution) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        execution,
        FormulaEnvironmentSpec.empty(),
        List.of());
  }

  private static AcceptancePolicy preservationPolicy() {
    return new AcceptancePolicy.RequireAll(
        List.of(
            new AcceptancePolicy.Requirement.PreserveInspectionFacts(
                List.of(
                    new InspectionStep(
                        new StepId("preserve-cell"),
                        new CellSelector.ByAddress(new SheetName("Report"), new CellAddress("B2")),
                        new SheetIntrospectionQuery.GetCells())))));
  }

  private static AcceptancePolicy opaquePartPolicy() {
    return new AcceptancePolicy.RequireAll(
        List.of(
            new AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts(
                List.of("/customXml/item1.xml"))));
  }
}
