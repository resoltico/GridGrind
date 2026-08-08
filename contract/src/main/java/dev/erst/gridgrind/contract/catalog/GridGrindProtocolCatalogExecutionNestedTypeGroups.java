package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;

import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import java.util.List;

/** Owns execution-mode nested type groups published by the protocol catalog. */
final class GridGrindProtocolCatalogExecutionNestedTypeGroups {
  private GridGrindProtocolCatalogExecutionNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> EXECUTION_GROUPS =
      List.of(
          nestedTypeGroup(
              "executionModeTypes",
              ExecutionModeInput.class,
              List.of(
                  descriptor(
                      ExecutionModeInput.FullXssf.class,
                      "FULL_XSSF",
                      GridGrindExecutionModeMetadata.fullXssf().catalogSummary()),
                  descriptor(
                      ExecutionModeInput.EventRead.class,
                      "EVENT_READ",
                      GridGrindExecutionModeMetadata.eventRead().catalogSummary()),
                  descriptor(
                      ExecutionModeInput.StreamingWrite.class,
                      "STREAMING_WRITE",
                      GridGrindExecutionModeMetadata.streamingWrite().catalogSummary()))),
          nestedTypeGroup(
              "executionJournalPhaseTypes",
              ExecutionJournal.Phase.class,
              List.of(
                  descriptor(
                      ExecutionJournal.Phase.NotStarted.class,
                      "NOT_STARTED",
                      "Execution phase that never began and therefore exposes no timing payload."),
                  descriptor(
                      ExecutionJournal.Phase.NotRequested.class,
                      "NOT_REQUESTED",
                      "Execution phase intentionally skipped and therefore exposing no timing payload."),
                  descriptor(
                      ExecutionJournal.Phase.Succeeded.class,
                      "SUCCEEDED",
                      "Execution phase that completed successfully and may include one timing payload."),
                  descriptor(
                      ExecutionJournal.Phase.Failed.class,
                      "FAILED",
                      "Execution phase that ended in failure and may include one timing payload."))),
          nestedTypeGroup(
              "executionJournalOutcomeTypes",
              ExecutionJournal.Outcome.class,
              List.of(
                  descriptor(
                      ExecutionJournal.Outcome.Succeeded.class,
                      "SUCCEEDED",
                      "Execution outcome summary for a successful run."),
                  descriptor(
                      ExecutionJournal.Outcome.Failed.class,
                      "FAILED",
                      "Execution outcome summary for a failed run, including the canonical failure"
                          + " code and an optional failing-step reference."))),
          nestedTypeGroup(
              "ooxmlEncryptionReportTypes",
              OoxmlEncryptionReport.class,
              List.of(
                  descriptor(
                      OoxmlEncryptionReport.None.class,
                      "NONE",
                      "Package carries no OOXML encryption envelope."),
                  descriptor(
                      OoxmlEncryptionReport.Encrypted.class,
                      "ENCRYPTED",
                      "Package is encrypted with one fully specified OOXML encryption envelope."))));
}
