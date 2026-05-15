package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
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
                      GridGrindExecutionModeMetadata.streamingWrite().catalogSummary()))));
}
