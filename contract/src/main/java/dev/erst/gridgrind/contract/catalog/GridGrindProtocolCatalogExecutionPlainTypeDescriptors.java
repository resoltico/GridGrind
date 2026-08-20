package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import java.util.List;

/** Execution/runtime and package-security plain type descriptors. */
final class GridGrindProtocolCatalogExecutionPlainTypeDescriptors {
  private GridGrindProtocolCatalogExecutionPlainTypeDescriptors() {}

  static final List<CatalogPlainTypeDescriptor> DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "executionJournalType",
              ExecutionJournal.class,
              "ExecutionJournal",
              "Structured execution telemetry returned on every success and failure response,"
                  + " including validation, input resolution, open, calculation, step,"
                  + " persistence, and close phases."),
          plainTypeDescriptor(
              "executionJournalSourceSummaryType",
              ExecutionJournal.SourceSummary.class,
              "ExecutionJournalSourceSummary",
              "Journal summary of the authored workbook source."),
          plainTypeDescriptor(
              "executionJournalStepType",
              ExecutionJournal.Step.class,
              "ExecutionJournalStep",
              "Per-step execution telemetry with resolved targets, timing, outcome,"
                  + " and optional failure detail."),
          plainTypeDescriptor(
              "executionJournalTargetType",
              ExecutionJournal.Target.class,
              "ExecutionJournalTarget",
              "One canonical target label recorded inside a step journal."),
          plainTypeDescriptor(
              "executionJournalFailureClassificationType",
              ExecutionJournal.FailureClassification.class,
              "ExecutionJournalFailureClassification",
              "Structured problem-code classification for one failed step."),
          plainTypeDescriptor(
              "executionJournalCalculationType",
              ExecutionJournal.Calculation.class,
              "ExecutionJournalCalculation",
              "Top-level calculation preflight and execution timings for one request."),
          plainTypeDescriptor(
              "executionJournalTimingType",
              ExecutionJournal.Timing.class,
              "ExecutionJournalTiming",
              "Measured timestamps and duration for one execution phase that actually ran."),
          plainTypeDescriptor(
              "executionJournalFailureStepType",
              ExecutionJournal.FailureStep.class,
              "ExecutionJournalFailureStep",
              "Canonical failing-step reference recorded when an execution failure is attributable"
                  + " to one authored step."),
          plainTypeDescriptor(
              "executionJournalEventType",
              ExecutionJournal.Event.class,
              "ExecutionJournalEvent",
              "Fine-grained verbose execution event emitted for live CLI rendering."),
          plainTypeDescriptor(
              "requestWarningType",
              RequestWarning.class,
              "RequestWarning",
              "Non-fatal authored-plan warning surfaced at the response root."),
          plainTypeDescriptor(
              "executionPolicyInputType",
              ExecutionPolicyInput.class,
              "ExecutionPolicyInput",
              GridGrindContractText.executionPolicyInputSummary()),
          plainTypeDescriptor(
              "calculationPolicyInputType",
              CalculationPolicyInput.class,
              "CalculationPolicyInput",
              GridGrindContractText.calculationPolicyInputSummary()),
          plainTypeDescriptor(
              "executionJournalInputType",
              ExecutionJournalInput.class,
              "ExecutionJournalInput",
              GridGrindContractText.executionJournalInputSummary()),
          plainTypeDescriptor(
              "calculationReportType",
              CalculationReport.class,
              "CalculationReport",
              "Structured calculation policy, preflight classification, and execution outcome"
                  + " returned on every success and failure response."),
          plainTypeDescriptor(
              "calculationPreflightType",
              CalculationReport.Preflight.class,
              "CalculationPreflightReport",
              "Formula capability classification captured before server-side evaluation begins."),
          plainTypeDescriptor(
              "calculationSummaryType",
              CalculationReport.Summary.class,
              "CalculationPreflightSummary",
              "Aggregate counts for evaluable, unevaluable, and unparseable formulas."),
          plainTypeDescriptor(
              "formulaCapabilityType",
              CalculationReport.FormulaCapability.class,
              "FormulaCapabilityReport",
              "One classified formula-cell capability entry from calculation preflight."),
          plainTypeDescriptor(
              "calculationExecutionType",
              CalculationReport.Execution.class,
              "CalculationExecutionReport",
              "Post-execution outcome for the authored calculation policy."),
          plainTypeDescriptor(
              "formulaEnvironmentInputType",
              FormulaEnvironmentInput.class,
              "FormulaEnvironmentInput",
              GridGrindContractText.formulaEnvironmentInputSummary()),
          plainTypeDescriptor(
              "ooxmlOpenSecurityInputType",
              OoxmlOpenSecurityInput.class,
              "OoxmlOpenSecurityInput",
              "Optional OOXML package-open settings for encrypted existing workbook sources."
                  + " password unlocks the encrypted OOXML package before GridGrind opens the"
                  + " inner .xlsx workbook."),
          plainTypeDescriptor(
              "ooxmlPersistenceSecurityInputType",
              OoxmlPersistenceSecurityInput.class,
              "OoxmlPersistenceSecurityInput",
              "Total OOXML encryption and signature policy applied during persistence."
                  + " It is required for writing EXISTING sources; NEW writes default to"
                  + " encryption NONE and signature NONE when security is omitted."),
          plainTypeDescriptor(
              "ooxmlEncryptionInputType",
              OoxmlEncryptionInput.class,
              "OoxmlEncryptionInput",
              GridGrindOoxmlWriteEncryptionContractText.inputSummary()),
          plainTypeDescriptorWithNotes(
              "ooxmlSignatureInputType",
              OoxmlSignatureInput.class,
              "OoxmlSignatureInput",
              "OOXML package-signing settings for workbook persistence."
                  + " keyPassword defaults to keystorePassword and digestAlgorithm defaults to"
                  + " SHA256 when omitted."
                  + " alias may be omitted only when the keystore contains exactly one"
                  + " signable private-key entry.",
              GridGrindProtocolCatalogNotes.requestOwnedPathRuleRef()),
          plainTypeDescriptor(
              "ooxmlPackageSecurityReportType",
              OoxmlPackageSecurityReport.class,
              "OoxmlPackageSecurityReport",
              "Factual OOXML package-security report covering encryption and package signatures."),
          plainTypeDescriptor(
              "ooxmlSignatureReportType",
              OoxmlSignatureReport.class,
              "OoxmlSignatureReport",
              "Factual OOXML package-signature report for one signature part."
                  + " state reflects the currently loaded workbook package, including"
                  + " INVALIDATED_BY_MUTATION for source signatures after in-memory edits."),
          plainTypeDescriptor(
              "ooxmlSignatureSignerIdentityType",
              OoxmlSignatureReport.SignerIdentity.class,
              "OoxmlSignatureSignerIdentity",
              "Signer identity material attached to one OOXML package signature report."),
          plainTypeDescriptorWithNotes(
              "formulaExternalWorkbookInputType",
              FormulaExternalWorkbookInput.class,
              "FormulaExternalWorkbookInput",
              "One external workbook binding keyed by the workbook name used inside formulas."
                  + " Workbook paths resolve through the shared request-owned path rule.",
              GridGrindProtocolCatalogNotes.requestOwnedPathRuleRef()),
          plainTypeDescriptor(
              "formulaUdfToolpackInputType",
              FormulaUdfToolpackInput.class,
              "FormulaUdfToolpackInput",
              "One named collection of template-backed user-defined functions."),
          plainTypeDescriptor(
              "formulaUdfFunctionInputType",
              FormulaUdfFunctionInput.class,
              "FormulaUdfFunctionInput",
              "One template-backed user-defined function."
                  + " formulaTemplate may reference ARG1, ARG2, and higher placeholders."
                  + " maximumArgumentCount defaults to minimumArgumentCount when omitted."));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group, Class<? extends Record> recordType, String id, String summary) {
    return CatalogTypeEntryFactory.plainTypeDescriptor(group, recordType, id, summary);
  }

  private static CatalogPlainTypeDescriptor plainTypeDescriptorWithNotes(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> noteRefs) {
    return CatalogTypeEntryFactory.plainTypeDescriptorWithNotes(
        group, recordType, id, summary, noteRefs);
  }
}
