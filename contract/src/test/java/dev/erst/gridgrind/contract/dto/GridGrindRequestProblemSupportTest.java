package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.json.ActionableInvariantMessage;
import dev.erst.gridgrind.contract.json.ActionableShapeMessage;
import dev.erst.gridgrind.contract.json.DuplicateStepId;
import dev.erst.gridgrind.contract.json.ExplicitNullField;
import dev.erst.gridgrind.contract.json.FieldValidationAddressRule;
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
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for typed request-problem wording and resolutions. */
class GridGrindRequestProblemSupportTest {
  @Test
  void canonicalMessageHelpersOwnMissingExplicitNullAndExtensionWording() {
    assertEquals(
        "Missing required field 'protocolVersion'",
        GridGrindRequestProblemSupport.missingRequiredFieldMessage("protocolVersion"));
    assertEquals(
        "Field 'planId' must be omitted when absent; explicit null is not accepted.",
        GridGrindRequestProblemSupport.explicitNullFieldMessage("planId"));
    assertEquals(
        "path must end in .xlsx (got: '.xlsm')",
        GridGrindRequestProblemSupport.message(
            new NonXlsxPath(".xlsm", java.util.Optional.empty())));
    assertEquals(
        "jsonPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindRequestProblemSupport.missingRequiredFieldMessage(" "))
            .getMessage());
  }

  @Test
  void messageRendersTypedSubtypeAndEnumProblemsWithoutReParsing() {
    assertEquals(
        "Unknown type value 'MOVEE'; use one of the published action type IDs; similar valid values: MOVE, MERGE",
        GridGrindRequestProblemSupport.message(
            new UnknownTypeValue(
                "MOVEE",
                java.util.Optional.of("steps[0].action.type"),
                java.util.List.of("MOVE", "MERGE"),
                java.util.Optional.of("use one of the published action type IDs"))));
    assertEquals(
        "Unknown type value 'NOPE'",
        GridGrindRequestProblemSupport.message(
            new UnknownTypeValue(
                "NOPE",
                java.util.Optional.empty(),
                java.util.List.of(),
                java.util.Optional.empty())));
    assertEquals(
        "Unsupported value 'BROKEN' for field 'steps[0].query.type'; expected one of: GET_CELL, GET_WINDOW",
        GridGrindRequestProblemSupport.message(
            new UnsupportedValue(
                "BROKEN",
                java.util.Optional.of("steps[0].query.type"),
                java.util.List.of("GET_CELL", "GET_WINDOW"))));
    assertEquals(
        "Unknown field 'steps[0].query.extra'",
        GridGrindRequestProblemSupport.message(new UnknownField("steps[0].query.extra")));
    assertEquals(
        "Field 'formulaEnvironment' must be omitted when absent; explicit null is not accepted.",
        GridGrindRequestProblemSupport.message(new ExplicitNullField("formulaEnvironment")));
    assertEquals(
        Optional.of("steps[0].query.extra"), new UnknownField("steps[0].query.extra").jsonPath());
    assertEquals(
        Optional.of("formulaEnvironment"), new ExplicitNullField("formulaEnvironment").jsonPath());
    assertEquals(
        Optional.of("steps[1].stepId"), new DuplicateStepId("dup", "steps[1].stepId").jsonPath());
  }

  @Test
  void resolutionUsesTypedDescriptorsAndProblemContextPaths() {
    ProblemContext.ReadRequest requestContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps[0].action.type"));
    ProblemContext.ReadRequest targetContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps[0].target.type"));
    ProblemContext.ReadRequest stepContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps[0].stepId"));

    assertEquals(
        "Add protocolVersion: \"V2\" at the request root.",
        GridGrindRequestProblemSupport.resolution(
            new MissingRequiredField("protocolVersion"), requestContext));
    assertEquals(
        "Add the required type discriminator at 'steps[0].action.type'.",
        GridGrindRequestProblemSupport.resolution(
            new MissingRequiredField("steps[0].action.type"), requestContext));
    assertEquals(
        "Add the required type discriminator at 'steps[0].target.type'.",
        GridGrindRequestProblemSupport.resolution(
            new MissingTypeDiscriminator("type"), targetContext));
    assertEquals(
        "Add required field 'steps[0].stepId' to the request payload.",
        GridGrindRequestProblemSupport.resolution(new MissingRequiredField("stepId"), stepContext));
    assertEquals(
        "Replace field 'steps[0].action.type' with one supported type value. Use --print-protocol-catalog --lookup or --search when you need the allowed values.",
        GridGrindRequestProblemSupport.resolution(
            new UnknownTypeValue(
                "MOVEE",
                java.util.Optional.empty(),
                java.util.List.of("MOVE"),
                java.util.Optional.empty()),
            requestContext));
    assertEquals(
        "Make every stepId unique. Rename or remove the duplicate value 'dup'.",
        GridGrindRequestProblemSupport.resolution(
            new DuplicateStepId("dup", "steps[1].stepId"), requestContext));
    assertEquals(
        "Provide a path ending in .xlsx for field 'source.path'.",
        GridGrindRequestProblemSupport.resolution(
            new NonXlsxPath(".xls", java.util.Optional.of("path")),
            new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.standardInput(),
                ProblemContextRequestSurfaces.JsonLocation.pathOnly("source.path"))));
    assertEquals(
        "Remove or rename unexpected field 'steps[0].query.extra' so the request matches the protocol.",
        GridGrindRequestProblemSupport.resolution(
            new UnknownField("steps[0].query.extra"), requestContext));
    assertEquals(
        "Replace field 'steps[0].query.type' with one supported value. Use --print-protocol-catalog --lookup or --search when you need the allowed values.",
        GridGrindRequestProblemSupport.resolution(
            new UnsupportedValue(
                "BROKEN",
                java.util.Optional.of("steps[0].query.type"),
                java.util.List.of("GET_CELL", "GET_WINDOW")),
            requestContext));
    assertEquals(
        "Replace the unsupported value with one allowed by the protocol.",
        GridGrindRequestProblemSupport.resolution(
            new UnsupportedValue("BROKEN", java.util.Optional.empty(), java.util.List.of("A")),
            requestContext));
    assertEquals(
        "Fix field 'steps[0].action.type' so it matches the published request shape.",
        GridGrindRequestProblemSupport.resolution(
            new MessageShape(
                "JSON object is missing required fields or has the wrong shape",
                java.util.Optional.empty()),
            requestContext));
    assertEquals(
        "Fix field 'steps[0].action.type' so it satisfies the request contract.",
        GridGrindRequestProblemSupport.resolution(
            new MessageInvariant("planId must not be blank", java.util.Optional.empty()),
            requestContext));
  }

  @Test
  void fieldValidationProblemsOwnCauseSpecificMessagesAndResolutions() {
    ProblemContext.ReadRequest unavailableContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.unavailable());

    assertEquals(
        "address must be a single-cell A1-style address",
        GridGrindRequestProblemSupport.message(
            FieldValidationProblem.atField("address", FieldValidationAddressRule.ADDRESS_SYNTAX)));
    assertEquals(
        "Use a single-cell A1-style address such as A1 or BC12 within Excel .xlsx bounds for field 'address'.",
        GridGrindRequestProblemSupport.resolution(
            FieldValidationProblem.atField("address", FieldValidationAddressRule.ADDRESS_SYNTAX),
            unavailableContext));
    assertEquals(
        "zoomPercent must be between 10 and 400 inclusive: 401",
        GridGrindRequestProblemSupport.message(
            FieldValidationProblem.atField(
                "zoomPercent", FieldValidationLayoutRule.ZOOM_PERCENT_RANGE, "10", "400", "401")));
    assertEquals(
        "Provide a zoom percentage between 10 and 400 inclusive for field 'zoomPercent'.",
        GridGrindRequestProblemSupport.resolution(
            FieldValidationProblem.atField(
                "zoomPercent", FieldValidationLayoutRule.ZOOM_PERCENT_RANGE, "10", "400", "401"),
            unavailableContext));
  }

  @Test
  void actionableDescriptorsPassThroughOwnedMessagesAndResolutions() {
    ProblemContext.ReadRequest requestContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps[0].target.address"));

    assertEquals(
        "Field 'target' must be a JSON object",
        GridGrindRequestProblemSupport.message(
            new ActionableShapeMessage(
                "Field 'target' must be a JSON object",
                "Replace the target selector with a JSON object.",
                Optional.empty())));
    assertEquals(
        "Replace the target selector with a JSON object.",
        GridGrindRequestProblemSupport.resolution(
            new ActionableShapeMessage(
                "Field 'target' must be a JSON object",
                "Replace the target selector with a JSON object.",
                Optional.empty()),
            requestContext));
    assertEquals(
        "address must identify one cell",
        GridGrindRequestProblemSupport.message(
            new ActionableInvariantMessage(
                "address must identify one cell",
                "Use a single-cell A1-style address for field 'address'.",
                Optional.of("address"))));
    assertEquals(
        "Use a single-cell A1-style address for field 'address'.",
        GridGrindRequestProblemSupport.resolution(
            new ActionableInvariantMessage(
                "address must identify one cell",
                "Use a single-cell A1-style address for field 'address'.",
                Optional.of("address")),
            requestContext));
  }

  @Test
  void indexedContextPathsFlowThroughPublicResolutions() {
    ProblemContext.ReadRequest indexedContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps[0]"));
    ProblemContext.ReadRequest parentContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps"));
    ProblemContext.ReadRequest unavailableContext =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.unavailable());

    assertEquals(
        "Remove or rename unexpected field 'steps[0]' so the request matches the protocol.",
        GridGrindRequestProblemSupport.resolution(new UnknownField("[0]"), indexedContext));
    assertEquals(
        "Remove or rename unexpected field '[0]' so the request matches the protocol.",
        GridGrindRequestProblemSupport.resolution(new UnknownField("[0]"), parentContext));
    assertEquals(
        "Add the required type discriminator at 'type'.",
        GridGrindRequestProblemSupport.resolution(
            new MissingTypeDiscriminator("type"), unavailableContext));
    assertEquals(
        "Provide a path ending in .xlsx for field 'path'.",
        GridGrindRequestProblemSupport.resolution(
            new NonXlsxPath(".xls", java.util.Optional.of("path")), unavailableContext));
  }
}
