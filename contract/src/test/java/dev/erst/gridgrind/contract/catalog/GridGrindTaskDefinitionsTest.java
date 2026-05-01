package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
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
  void builtInStarterTemplatesDoNotShipBrokenLookupsOrTodoPlaceholders() {
    assertDoesNotThrow(() -> GridGrindTaskPlanner.templateFor("CUSTOM_XML_WORKFLOW"));
    assertDoesNotThrow(
        () ->
            GridGrindTaskDefinitions.definitions().stream()
                .map(TaskDefinition::starterTemplate)
                .forEach(
                    template ->
                        assertTrue(
                            template.authoringNotes().stream()
                                .noneMatch(note -> note.contains("TODO_")))));
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

  @Test
  void starterTemplateValidatorsRejectTodoPlaceholdersAndBrokenLookupNotes() throws IOException {
    IllegalStateException requestTodoFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindTaskDefinitions.validateStarterTemplatePlaceholders(
                    taskDefinition(
                        "BROKEN_REQUEST",
                        request(
                            """
                            {
                              "protocolVersion": "V1",
                              "source": { "type": "NEW" },
                              "persistence": { "type": "NONE" },
                              "execution": {
                                "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
                                "journal": { "level": "NORMAL" },
                                "calculation": {
                                  "strategy": { "type": "DO_NOT_CALCULATE" },
                                  "markRecalculateOnOpen": false
                                }
                              },
                              "formulaEnvironment": {
                                "externalWorkbooks": [],
                                "missingWorkbookPolicy": "ERROR",
                                "udfToolpacks": []
                              },
                              "steps": [
                                {
                                  "stepId": "TODO_STEP",
                                  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                                  "action": { "type": "ENSURE_SHEET" }
                                }
                              ]
                            }
                            """),
                        "Clean note")));
    assertTrue(requestTodoFailure.getMessage().contains("starter request template"));

    IllegalStateException noteTodoFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindTaskDefinitions.validateStarterTemplatePlaceholders(
                    taskDefinition(
                        "BROKEN_NOTE",
                        GridGrindProtocolCatalog.requestTemplate(),
                        "Remember TODO_PROTOCOL_LOOKUP before publishing.")));
    assertTrue(noteTodoFailure.getMessage().contains("authoring note"));

    IllegalStateException brokenLookupFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindTaskDefinitions.validateStarterTemplateLookups(
                    taskDefinition(
                        "BROKEN_LOOKUP",
                        GridGrindProtocolCatalog.requestTemplate(),
                        "No protocol lookup in this first note.",
                        "Use --print-protocol-catalog --operation mutationActionTypes:NO_SUCH_ACTION.")));
    assertTrue(brokenLookupFailure.getMessage().contains("unknown protocol lookup"));
  }

  @Test
  void starterTemplateLookupValidatorAcceptsKnownLookupsWithTrailingPunctuation() {
    assertDoesNotThrow(
        () ->
            GridGrindTaskDefinitions.validateStarterTemplateLookups(
                taskDefinition(
                    "VALID_LOOKUPS",
                    GridGrindProtocolCatalog.requestTemplate(),
                    "Use --print-protocol-catalog --operation mutationActionTypes:SET_CELL,"
                        + " then --print-protocol-catalog --operation assertionTypes:EXPECT_CELL_VALUE).")));
  }

  private static TaskDefinition taskDefinition(
      String id, WorkbookPlan requestTemplate, String... authoringNotes) {
    return new TaskDefinition(
        new TaskEntry(
            id,
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")),
                    List.of("note"))),
            List.of("pitfall")),
        List.of("search"),
        requestTemplate,
        List.of(authoringNotes));
  }

  private static WorkbookPlan request(String json) throws IOException {
    return GridGrindJson.readRequest(json.getBytes(java.nio.charset.StandardCharsets.UTF_8));
  }
}
