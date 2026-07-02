package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.ActionableInvariantMessage;
import dev.erst.gridgrind.contract.json.ActionableShapeMessage;
import dev.erst.gridgrind.contract.json.DuplicateStepId;
import dev.erst.gridgrind.contract.json.ExplicitNullField;
import dev.erst.gridgrind.contract.json.FieldValidationLayoutRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.json.MessageInvariant;
import dev.erst.gridgrind.contract.json.MessageShape;
import dev.erst.gridgrind.contract.json.MissingRequiredField;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
import dev.erst.gridgrind.contract.json.NonXlsxPath;
import dev.erst.gridgrind.contract.json.UnknownField;
import dev.erst.gridgrind.contract.json.UnknownTypeValue;
import dev.erst.gridgrind.contract.json.UnsupportedValue;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ObjectNode;

/** Focused coverage for step-level doctor preflight batching and structured path rebasing. */
class CliDoctorRequestStepPreflightTest {
  private static final JsonMapper JSON = JsonMapper.builder().build();

  @Test
  void fromReturnsCleanResultWhenStepsArrayIsMissing() {
    CliDoctorRequestStepPreflight.StepPreflight preflight =
        CliDoctorRequestStepPreflight.from(
            JSON.createObjectNode(), ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(java.util.List.of(), preflight.problems());
    assertFalse(preflight.usesSyntheticValues());
    assertTrue(preflight.summaryTrustworthy());
  }

  @Test
  void fromReturnsCleanResultWhenStepsArrayContainsOnlyValidSteps() throws IOException {
    CliDoctorRequestStepPreflight.StepPreflight preflight =
        CliDoctorRequestStepPreflight.from(
            objectNode(
                """
                {
                  "steps": [
                    {
                      "stepId": "summary",
                      "target": { "type": "WORKBOOK_CURRENT" },
                      "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    }
                  ]
                }
                """),
            ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(java.util.List.of(), preflight.problems());
    assertFalse(preflight.usesSyntheticValues());
    assertTrue(preflight.summaryTrustworthy());
  }

  @Test
  void fromRebasesMalformedStepShapeProblemsToTheAuthoredIndex() throws IOException {
    CliDoctorRequestStepPreflight.StepPreflight preflight =
        CliDoctorRequestStepPreflight.from(
            objectNode(
                """
                {
                  "steps": [
                    {
                      "stepId": "summary",
                      "target": { "type": "WORKBOOK_CURRENT" },
                      "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    },
                    {
                      "stepId": "broken-query",
                      "target": { "type": "WORKBOOK_CURRENT" },
                      "query": { }
                    }
                  ]
                }
                """),
            ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"));

    ProblemContext.ReadRequest context =
        assertInstanceOf(
            ProblemContext.ReadRequest.class, preflight.problems().getFirst().context());

    assertEquals(1, preflight.problems().size());
    assertEquals(java.util.Optional.of("steps[1].query.type"), context.jsonPath());
    assertEquals(
        "Missing required field 'steps[1].query.type'", preflight.problems().getFirst().message());
    assertTrue(preflight.usesSyntheticValues());
    assertFalse(preflight.summaryTrustworthy());
  }

  @Test
  void fromRebasesMalformedStepInvariantProblemsToTheAuthoredIndex() throws IOException {
    CliDoctorRequestStepPreflight.StepPreflight preflight =
        CliDoctorRequestStepPreflight.from(
            objectNode(
                """
                {
                  "steps": [
                    {
                      "stepId": "summary",
                      "target": { "type": "WORKBOOK_CURRENT" },
                      "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    },
                    {
                      "stepId": "bad-sheet",
                      "target": { "type": "SHEET_BY_NAME", "name": "Bad:Name" },
                      "action": { "type": "ENSURE_SHEET" }
                    }
                  ]
                }
                """),
            ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"));

    ProblemContext.ReadRequest context =
        assertInstanceOf(
            ProblemContext.ReadRequest.class, preflight.problems().getFirst().context());

    assertEquals(1, preflight.problems().size());
    assertEquals(Optional.of("steps[1].target.name"), context.jsonPath());
    assertTrue(preflight.problems().getFirst().message().contains("invalid Excel character ':'"));
    assertTrue(preflight.usesSyntheticValues());
    assertFalse(preflight.summaryTrustworthy());
  }

  @Test
  void rebaseStepProblemLeavesStepZeroDescriptorUntouched() {
    MissingRequiredField rebased =
        assertInstanceOf(
            MissingRequiredField.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new MissingRequiredField("steps[0].query.type"), 0));

    assertEquals("steps[0].query.type", rebased.jsonPathValue());
  }

  @Test
  void rebaseStepProblemLeavesNonStepPathsUntouched() {
    MissingRequiredField rebased =
        assertInstanceOf(
            MissingRequiredField.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new MissingRequiredField("query.type"), 7));

    assertEquals("query.type", rebased.jsonPathValue());
  }

  @Test
  void rebaseStepProblemRewritesStructuredPathsForLaterIndexes() {
    UnsupportedValue rebased =
        assertInstanceOf(
            UnsupportedValue.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new UnsupportedValue(
                    "BROKEN",
                    java.util.Optional.of("steps[0].query.type"),
                    java.util.List.of("GET_WINDOW")),
                2));

    assertEquals(java.util.Optional.of("steps[2].query.type"), rebased.jsonPath());
    assertEquals(
        "Unsupported value 'BROKEN' for field 'steps[2].query.type'; expected one of: GET_WINDOW",
        dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport.message(rebased));
  }

  @Test
  void rebaseStepProblemCoversRemainingDescriptorFamilies() {
    MissingTypeDiscriminator missingType =
        assertInstanceOf(
            MissingTypeDiscriminator.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new MissingTypeDiscriminator("steps[0].query.type"), 4));
    UnknownField unknownField =
        assertInstanceOf(
            UnknownField.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new UnknownField("steps[0].query.extra"), 4));
    UnknownTypeValue unknownType =
        assertInstanceOf(
            UnknownTypeValue.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new UnknownTypeValue(
                    "BROKEN",
                    Optional.of("steps[0].query.type"),
                    java.util.List.of("GET_WORKBOOK_SUMMARY"),
                    Optional.of("use a valid query type")),
                4));
    ExplicitNullField explicitNull =
        assertInstanceOf(
            ExplicitNullField.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new ExplicitNullField("steps[0].query"), 4));
    MessageShape messageShape =
        assertInstanceOf(
            MessageShape.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new MessageShape("wrong shape", Optional.of("steps[0].query")), 4));
    ActionableShapeMessage actionableShape =
        assertInstanceOf(
            ActionableShapeMessage.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new ActionableShapeMessage(
                    "wrong shape", "repair the shape", Optional.of("steps[0].query")),
                4));
    DuplicateStepId duplicateStepId =
        assertInstanceOf(
            DuplicateStepId.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new DuplicateStepId("dup", "steps[0].stepId"), 4));
    NonXlsxPath nonXlsxPath =
        assertInstanceOf(
            NonXlsxPath.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new NonXlsxPath(".xlsm", Optional.of("steps[0].persistence.path")), 4));
    MessageInvariant messageInvariant =
        assertInstanceOf(
            MessageInvariant.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new MessageInvariant("bad invariant", Optional.of("steps[0].query")), 4));
    ActionableInvariantMessage actionableInvariant =
        assertInstanceOf(
            ActionableInvariantMessage.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new ActionableInvariantMessage(
                    "bad invariant", "repair the invariant", Optional.of("steps[0].query")),
                4));

    assertEquals("steps[4].query.type", missingType.jsonPathValue());
    assertEquals("steps[4].query.extra", unknownField.jsonPathValue());
    assertEquals(Optional.of("steps[4].query.type"), unknownType.jsonPath());
    assertEquals("steps[4].query", explicitNull.jsonPathValue());
    assertEquals(Optional.of("steps[4].query"), messageShape.jsonPath());
    assertEquals(Optional.of("steps[4].query"), actionableShape.jsonPath());
    assertEquals("steps[4].stepId", duplicateStepId.jsonPathValue());
    assertEquals(Optional.of("steps[4].persistence.path"), nonXlsxPath.jsonPath());
    assertEquals(Optional.of("steps[4].query"), messageInvariant.jsonPath());
    assertEquals(Optional.of("steps[4].query"), actionableInvariant.jsonPath());
  }

  @Test
  void rebaseStepProblemLeavesOwnedDescriptorTextUnchangedWhileRebasingJsonPath() {
    ActionableShapeMessage actionableShape =
        assertInstanceOf(
            ActionableShapeMessage.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new ActionableShapeMessage(
                    "Field 'steps[0].query.type' must be a string",
                    "Replace field 'steps[0].query.type' with a JSON string type id.",
                    Optional.of("steps[0].query.type")),
                4));
    MessageShape messageShape =
        assertInstanceOf(
            MessageShape.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new MessageShape(
                    "Fix field 'steps[0].query.type' before retrying.",
                    Optional.of("steps[0].query.type")),
                4));
    ActionableInvariantMessage actionableInvariant =
        assertInstanceOf(
            ActionableInvariantMessage.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new ActionableInvariantMessage(
                    "Field 'steps[0].query.type' is invalid",
                    "Repair field 'steps[0].query.type'.",
                    Optional.of("steps[0].query.type")),
                4));
    MessageInvariant messageInvariant =
        assertInstanceOf(
            MessageInvariant.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new MessageInvariant(
                    "Field 'steps[0].query.type' is inconsistent",
                    Optional.of("steps[0].query.type")),
                4));

    assertEquals("Field 'steps[0].query.type' must be a string", actionableShape.message());
    assertEquals(
        "Replace field 'steps[0].query.type' with a JSON string type id.",
        actionableShape.resolutionValue());
    assertEquals("Fix field 'steps[0].query.type' before retrying.", messageShape.message());
    assertEquals("Field 'steps[0].query.type' is invalid", actionableInvariant.message());
    assertEquals("Repair field 'steps[0].query.type'.", actionableInvariant.resolutionValue());
    assertEquals("Field 'steps[0].query.type' is inconsistent", messageInvariant.message());
    assertEquals(Optional.of("steps[4].query.type"), actionableShape.jsonPath());
    assertEquals(Optional.of("steps[4].query.type"), messageShape.jsonPath());
    assertEquals(Optional.of("steps[4].query.type"), actionableInvariant.jsonPath());
    assertEquals(Optional.of("steps[4].query.type"), messageInvariant.jsonPath());
  }

  @Test
  void rebaseStepProblemLeavesFieldOwnedActionableTextLocal() {
    ActionableShapeMessage actionableShape =
        assertInstanceOf(
            ActionableShapeMessage.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new ActionableShapeMessage(
                    "Field 'target.type' must be a string",
                    "Replace field 'target.type' with a JSON string selector id.",
                    Optional.of("steps[0].target.type")),
                4));

    assertEquals("Field 'target.type' must be a string", actionableShape.message());
    assertEquals(
        "Replace field 'target.type' with a JSON string selector id.",
        actionableShape.resolutionValue());
    assertEquals(Optional.of("steps[4].target.type"), actionableShape.jsonPath());
  }

  @Test
  void rebaseStepProblemLeavesPathlessActionableDescriptorsUntouched() {
    ActionableShapeMessage actionableShape =
        assertInstanceOf(
            ActionableShapeMessage.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new ActionableShapeMessage("wrong shape", "repair the shape", Optional.empty()),
                4));

    assertEquals("wrong shape", actionableShape.message());
    assertEquals("repair the shape", actionableShape.resolutionValue());
    assertEquals(Optional.empty(), actionableShape.jsonPath());
  }

  @Test
  void rebaseStepProblemRebasesStructuredFieldValidationPaths() {
    FieldValidationProblem rebased =
        assertInstanceOf(
            FieldValidationProblem.class,
            CliDoctorRequestStepPreflight.rebaseStepProblem(
                new FieldValidationProblem(
                    "zoomPercent",
                    Optional.of("steps[0].action.zoomPercent"),
                    FieldValidationLayoutRule.ZOOM_PERCENT_RANGE,
                    List.of("10", "400", "401")),
                4));

    assertEquals(Optional.of("steps[4].action.zoomPercent"), rebased.jsonPath());
    assertEquals("zoomPercent", rebased.fieldName());
    assertEquals(
        "zoomPercent must be between 10 and 400 inclusive: 401",
        dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport.message(rebased));
  }

  private static ObjectNode objectNode(String json) throws IOException {
    JsonNode node = JSON.readTree(json);
    return (ObjectNode) node;
  }
}
