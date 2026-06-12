package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for request-problem path extraction and resolution wording. */
class GridGrindRequestProblemSupportTest {
  @Test
  void extractsActionableJsonPathsFromPublicRequestMessages() {
    assertEquals(
        Optional.of("protocolVersion"),
        GridGrindRequestProblemSupport.jsonPathFromMessage(
            "Missing required field 'protocolVersion'"));
    assertEquals(
        Optional.of("steps[0].action"),
        GridGrindRequestProblemSupport.jsonPathFromMessage("Unknown field 'steps[0].action'"));
    assertEquals(
        Optional.of("steps[0].action.type"),
        GridGrindRequestProblemSupport.jsonPathFromMessage(
            "Unsupported value 'MOVEE' for field 'steps[0].action.type'"));
    assertEquals(
        Optional.of("planId"),
        GridGrindRequestProblemSupport.jsonPathFromMessage("planId must not be blank"));
    assertEquals(Optional.empty(), GridGrindRequestProblemSupport.jsonPathFromMessage(null));
    assertEquals(
        Optional.empty(),
        GridGrindRequestProblemSupport.jsonPathFromMessage("Unknown type value 'MOVEE'"));
  }

  @Test
  void emitsSpecificResolutionsForEverySupportedProblemPattern() {
    assertEquals(
        Optional.of("Make every stepId unique. Rename or remove the duplicate value 'budget'."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST,
            "steps must not contain duplicate stepId values: budget",
            Optional.empty()));
    assertEquals(
        Optional.of("Add protocolVersion: \"V1\" at the request root."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST,
            "Missing required field 'protocolVersion'",
            Optional.empty()));
    assertEquals(
        Optional.of("Add the required type discriminator at 'steps[0].action.type'."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST,
            "Missing required field 'steps[0].action.type'",
            Optional.empty()));
    assertEquals(
        Optional.of(
            "Remove or rename unexpected field 'steps[0].bogus' so the request matches the protocol."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST,
            "Unknown field 'steps[0].bogus'",
            Optional.empty()));
    assertEquals(
        Optional.of(
            "Replace field 'steps[0].action.type' with one supported value. Use --print-protocol-catalog --lookup or --search when you need the allowed values."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST,
            "Unsupported value 'MOVEE' for field 'steps[0].action.type'",
            Optional.empty()));
    assertEquals(
        Optional.of("Provide a non-blank value for field 'planId'."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST, "planId must not be blank", Optional.empty()));
  }

  @Test
  void invalidRequestShapeResolutionsUseProblemContextJsonPaths() {
    ProblemContext.ReadRequest context =
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps[0].action.type"));

    assertEquals(
        Optional.of(
            "Replace field 'steps[0].action.type' with one supported type value. Use --print-protocol-catalog --lookup or --search when you need the allowed values."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE, "Unknown type value 'MOVEE'", context));
    assertEquals(
        Optional.of("Fix field 'steps[0].action.type' so it matches the published request shape."),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE, "Cannot deserialize value", context));
    assertEquals(
        Optional.empty(),
        GridGrindRequestProblemSupport.specificResolution(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE, " ", Optional.of("steps[0].action.type")));
  }
}
