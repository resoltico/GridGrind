package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ObjectNode;

/** Covers internal JSON-support branches that are otherwise hard to drive through one request. */
class GridGrindJsonSupportCoverageTest {
  @Test
  void mergedJsonPathKeepsExactAndSuffixMatchesButJoinsRelativeFields() {
    assertEquals(
        Optional.of("steps[0].target.name"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.of("steps[0].target.name"), Optional.of("steps[0].target.name")));
    assertEquals(
        Optional.of("steps[0].target.name"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.of("steps[0].target.name"), Optional.of("name")));
    assertEquals(
        Optional.of("steps[0]"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.of("steps[0]"), Optional.of("[0]")));
    assertEquals(
        Optional.of("steps[0].targetname.name"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.of("steps[0].targetname"), Optional.of("name")));
    assertEquals(
        Optional.of("steps[0].target.name"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.of("steps[0].target"), Optional.of("name")));
    assertEquals(
        Optional.of("query.name"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.of("steps[0].target"), Optional.of("query.name")));
    assertEquals(
        Optional.of("steps[0].query"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.empty(), Optional.of("steps[0].query")));
    assertEquals(
        Optional.of("steps[0].target"),
        GridGrindJsonProblemMessageSupport.mergedJsonPath(
            Optional.of("steps[0].target"), Optional.empty()));
  }

  @Test
  void relativeFieldDetectionRejectsNestedAndIndexedPaths() {
    assertTrue(GridGrindJsonProblemMessageSupport.isRelativeFieldPath("name"));
    assertFalse(GridGrindJsonProblemMessageSupport.isRelativeFieldPath("target.name"));
    assertFalse(GridGrindJsonProblemMessageSupport.isRelativeFieldPath("[0]"));
  }

  @Test
  void rawStepPayloadShapeValidationCoversNoOpAndConflictCases() {
    JsonMapper mapper = JsonMapper.builder().build();
    ObjectNode request = mapper.createObjectNode();

    assertDoesNotThrow(
        () ->
            GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(mapper.nullNode()));
    assertDoesNotThrow(
        () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));

    request.put("steps", 3);
    assertDoesNotThrow(
        () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));

    request.putArray("steps").add(3);
    assertDoesNotThrow(
        () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));

    request.set(
        "steps",
        mapper
            .createArrayNode()
            .add(mapper.createObjectNode().put("stepId", "empty").putObject("target")));
    InvalidRequestShapeException missingPayload =
        assertThrows(
            InvalidRequestShapeException.class,
            () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));
    assertEquals(Optional.of("steps[0]"), missingPayload.jsonPath());

    assertEquals(
        Optional.of("steps[0].assertion"),
        conflictingPayloadPath(
            mapper,
            """
            { "stepId": "both", "target": {}, "action": {}, "assertion": {} }
            """));
    assertEquals(
        Optional.of("steps[0].query"),
        conflictingPayloadPath(
            mapper,
            """
            { "stepId": "both", "target": {}, "action": {}, "query": {} }
            """));
    assertEquals(
        Optional.of("steps[0].query"),
        conflictingPayloadPath(
            mapper,
            """
            { "stepId": "both", "target": {}, "assertion": {}, "query": {} }
            """));
  }

  @Test
  void conflictingPayloadFieldRejectsImpossibleCallers() {
    JsonMapper mapper = JsonMapper.builder().build();
    ObjectNode emptyStep = mapper.createObjectNode();
    ObjectNode actionOnlyStep = mapper.createObjectNode();
    ObjectNode assertionOnlyStep = mapper.createObjectNode();
    actionOnlyStep.putObject("action");
    assertionOnlyStep.putObject("assertion");

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindJsonStepPayloadShapeSupport.conflictingPayloadField(emptyStep));
    assertThrows(
        IllegalStateException.class,
        () -> GridGrindJsonStepPayloadShapeSupport.conflictingPayloadField(actionOnlyStep));
    assertThrows(
        IllegalStateException.class,
        () -> GridGrindJsonStepPayloadShapeSupport.conflictingPayloadField(assertionOnlyStep));

    assertEquals("conflictingPayloadField requires at least two payloads", failure.getMessage());
  }

  private static Optional<String> conflictingPayloadPath(JsonMapper mapper, String stepJson) {
    ObjectNode request = mapper.createObjectNode();
    request.set(
        "steps", mapper.createArrayNode().add(assertDoesNotThrow(() -> mapper.readTree(stepJson))));
    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () -> GridGrindJsonStepPayloadShapeSupport.rejectInvalidStepPayloadShapes(request));
    return failure.jsonPath();
  }
}
