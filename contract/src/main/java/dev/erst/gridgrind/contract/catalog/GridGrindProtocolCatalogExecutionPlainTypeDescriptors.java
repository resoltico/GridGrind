package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
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
                  + " including validation, open, calculation, step, persistence, and close"
                  + " phases.",
              List.of(
                  "planId",
                  "level",
                  "source",
                  "persistence",
                  "validation",
                  "open",
                  "calculation",
                  "persistencePhase",
                  "close",
                  "steps",
                  "warnings",
                  "outcome",
                  "events")),
          plainTypeDescriptor(
              "executionJournalSourceSummaryType",
              ExecutionJournal.SourceSummary.class,
              "ExecutionJournalSourceSummary",
              "Journal summary of the authored workbook source.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPersistenceSummaryType",
              ExecutionJournal.PersistenceSummary.class,
              "ExecutionJournalPersistenceSummary",
              "Journal summary of the authored persistence policy.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPhaseType",
              ExecutionJournal.Phase.class,
              "ExecutionJournalPhase",
              "One timed execution phase with status, timestamps, and duration.",
              List.of("startedAt", "finishedAt")),
          plainTypeDescriptor(
              "executionJournalStepType",
              ExecutionJournal.Step.class,
              "ExecutionJournalStep",
              "Per-step execution telemetry with resolved targets, timing, outcome,"
                  + " and optional failure detail.",
              List.of("failure")),
          plainTypeDescriptor(
              "executionJournalTargetType",
              ExecutionJournal.Target.class,
              "ExecutionJournalTarget",
              "One canonical target label recorded inside a step journal.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalFailureClassificationType",
              ExecutionJournal.FailureClassification.class,
              "ExecutionJournalFailureClassification",
              "Structured problem-code classification for one failed step.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalCalculationType",
              ExecutionJournal.Calculation.class,
              "ExecutionJournalCalculation",
              "Top-level calculation preflight and execution timings for one request.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalOutcomeType",
              ExecutionJournal.Outcome.class,
              "ExecutionJournalOutcome",
              "Final outcome summary for one execution journal.",
              List.of("failedStepIndex", "failedStepId", "failureCode")),
          plainTypeDescriptor(
              "executionJournalEventType",
              ExecutionJournal.Event.class,
              "ExecutionJournalEvent",
              "Fine-grained verbose execution event emitted for live CLI rendering.",
              List.of("stepIndex", "stepId")),
          plainTypeDescriptor(
              "requestWarningType",
              RequestWarning.class,
              "RequestWarning",
              "Non-fatal authored-plan warning surfaced on success and echoed inside the execution journal.",
              List.of()),
          plainTypeDescriptor(
              "executionPolicyInputType",
              ExecutionPolicyInput.class,
              "ExecutionPolicyInput",
              GridGrindContractText.executionPolicyInputSummary(),
              List.of("mode", "journal", "calculation")),
          plainTypeDescriptor(
              "calculationPolicyInputType",
              CalculationPolicyInput.class,
              "CalculationPolicyInput",
              GridGrindContractText.calculationPolicyInputSummary(),
              List.of("strategy")),
          plainTypeDescriptor(
              "executionModeInputType",
              ExecutionModeInput.class,
              "ExecutionModeInput",
              GridGrindContractText.executionModeInputSummary(),
              List.of("readMode", "writeMode")),
          plainTypeDescriptor(
              "executionJournalInputType",
              ExecutionJournalInput.class,
              "ExecutionJournalInput",
              GridGrindContractText.executionJournalInputSummary(),
              List.of("level")),
          plainTypeDescriptor(
              "calculationReportType",
              CalculationReport.class,
              "CalculationReport",
              "Structured calculation policy, preflight classification, and execution outcome"
                  + " returned on every success and failure response.",
              List.of("preflight")),
          plainTypeDescriptor(
              "calculationPreflightType",
              CalculationReport.Preflight.class,
              "CalculationPreflightReport",
              "Formula capability classification captured before server-side evaluation begins.",
              List.of()),
          plainTypeDescriptor(
              "calculationSummaryType",
              CalculationReport.Summary.class,
              "CalculationPreflightSummary",
              "Aggregate counts for evaluable, unevaluable, and unparseable formulas.",
              List.of()),
          plainTypeDescriptor(
              "formulaCapabilityType",
              CalculationReport.FormulaCapability.class,
              "FormulaCapabilityReport",
              "One classified formula-cell capability entry from calculation preflight.",
              List.of("problemCode", "message")),
          plainTypeDescriptor(
              "calculationExecutionType",
              CalculationReport.Execution.class,
              "CalculationExecutionReport",
              "Post-execution outcome for the authored calculation policy.",
              List.of("message")),
          plainTypeDescriptor(
              "formulaEnvironmentInputType",
              FormulaEnvironmentInput.class,
              "FormulaEnvironmentInput",
              "Request-scoped formula-evaluation environment covering external workbook bindings,"
                  + " missing-workbook policy, and template-backed UDF toolpacks.",
              List.of("externalWorkbooks", "missingWorkbookPolicy", "udfToolpacks")),
          plainTypeDescriptor(
              "ooxmlOpenSecurityInputType",
              OoxmlOpenSecurityInput.class,
              "OoxmlOpenSecurityInput",
              "Optional OOXML package-open settings for encrypted existing workbook sources."
                  + " password unlocks the encrypted OOXML package before GridGrind opens the"
                  + " inner .xlsx workbook.",
              List.of("password")),
          plainTypeDescriptor(
              "ooxmlPersistenceSecurityInputType",
              OoxmlPersistenceSecurityInput.class,
              "OoxmlPersistenceSecurityInput",
              "Optional OOXML package-security settings applied during persistence."
                  + " Supply encryption, signature, or both.",
              List.of("encryption", "signature")),
          plainTypeDescriptor(
              "ooxmlEncryptionInputType",
              OoxmlEncryptionInput.class,
              "OoxmlEncryptionInput",
              "OOXML package-encryption settings for workbook persistence."
                  + " mode defaults to AGILE when omitted.",
              List.of("mode")),
          plainTypeDescriptor(
              "ooxmlSignatureInputType",
              OoxmlSignatureInput.class,
              "OoxmlSignatureInput",
              "OOXML package-signing settings for workbook persistence."
                  + " pkcs12Path follows the request-owned path rule."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " keyPassword defaults to keystorePassword and digestAlgorithm defaults to"
                  + " SHA256 when omitted."
                  + " alias may be omitted only when the keystore contains exactly one"
                  + " signable private-key entry.",
              List.of("keyPassword", "alias", "digestAlgorithm", "description")),
          plainTypeDescriptor(
              "ooxmlPackageSecurityReportType",
              OoxmlPackageSecurityReport.class,
              "OoxmlPackageSecurityReport",
              "Factual OOXML package-security report covering encryption and package signatures.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlEncryptionReportType",
              OoxmlEncryptionReport.class,
              "OoxmlEncryptionReport",
              "Factual OOXML package-encryption report for one workbook package."
                  + " Detail fields are present only when encrypted=true.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlSignatureReportType",
              OoxmlSignatureReport.class,
              "OoxmlSignatureReport",
              "Factual OOXML package-signature report for one signature part."
                  + " state reflects the currently loaded workbook package, including"
                  + " INVALIDATED_BY_MUTATION for source signatures after in-memory edits.",
              List.of()),
          plainTypeDescriptor(
              "formulaExternalWorkbookInputType",
              FormulaExternalWorkbookInput.class,
              "FormulaExternalWorkbookInput",
              "One external workbook binding keyed by the workbook name used inside formulas."
                  + " path follows the request-owned path rule."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary(),
              List.of()),
          plainTypeDescriptor(
              "formulaUdfToolpackInputType",
              FormulaUdfToolpackInput.class,
              "FormulaUdfToolpackInput",
              "One named collection of template-backed user-defined functions.",
              List.of()),
          plainTypeDescriptor(
              "formulaUdfFunctionInputType",
              FormulaUdfFunctionInput.class,
              "FormulaUdfFunctionInput",
              "One template-backed user-defined function."
                  + " formulaTemplate may reference ARG1, ARG2, and higher placeholders."
                  + " maximumArgumentCount defaults to minimumArgumentCount when omitted.",
              List.of("maximumArgumentCount")));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return GridGrindProtocolCatalog.plainTypeDescriptor(
        group, recordType, id, summary, optionalFields);
  }
}
