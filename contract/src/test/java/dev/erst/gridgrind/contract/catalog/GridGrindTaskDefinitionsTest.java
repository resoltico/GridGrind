package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage and guardrail tests for the internal task-definition registry. */
class GridGrindTaskDefinitionsTest {
  @Test
  void protocolLookupNotesRenderOxfordCommaListsAndRejectEmptyInputs() {
    assertEquals(
        "Use --print-protocol-catalog --operation mutationActionTypes:SET_CELL,"
            + " --print-protocol-catalog --operation mutationActionTypes:SET_RANGE, and"
            + " --print-protocol-catalog --operation mutationActionTypes:SET_TABLE for exact"
            + " workbook authoring shapes, or --print-protocol-catalog --search cells,"
            + " --print-protocol-catalog --search ranges, and"
            + " --print-protocol-catalog --search tables when browsing nearby surfaces.",
        GridGrindTaskDefinitionSupport.protocolLookupNote(
            "workbook authoring shapes",
            List.of(
                "mutationActionTypes:SET_CELL",
                "mutationActionTypes:SET_RANGE",
                "mutationActionTypes:SET_TABLE"),
            List.of("cells", "ranges", "tables")));

    IllegalArgumentException invocation =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                GridGrindTaskDefinitionSupport.protocolLookupNote(
                    "workbook authoring shapes", List.of(), List.of("cells")));

    assertEquals("values must not be empty", invocation.getMessage());
  }

  @Test
  void builtInTaskDefinitionsValidateAgainstTheProtocolCatalog() {
    assertDoesNotThrow(GridGrindTaskDefinitions::validateCapabilityReferences);
  }

  @Test
  void internalTaskDefinitionValidatorRejectsDanglingCapabilityReferences() {
    IllegalStateException invocation =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindTaskDefinitions.validateTaskCapabilityReferences(
                    List.of(
                        new TaskEntry(
                            "BROKEN_TASK",
                            "summary",
                            List.of("office"),
                            List.of("outcome"),
                            List.of("input"),
                            List.of("feature"),
                            List.of(
                                new TaskPhase(
                                    "Phase",
                                    "Objective",
                                    List.of(
                                        new TaskCapabilityRef(
                                            "inspectionQueryTypes", "NO_SUCH_QUERY")),
                                    List.of("note"))),
                            List.of("pitfall")))));

    assertTrue(
        invocation.getMessage().contains("BROKEN_TASK references unknown protocol capability"));
  }
}
