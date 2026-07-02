package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

/** Direct coverage for request-shape detection at the tree-to-record decode seam. */
class GridGrindJsonCodecSupportTest {
  @Test
  void decodeTreeRejectsExplicitNullMembersBeforeBinding() throws IOException {
    JsonNode requestTree =
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER.readTree(
            """
            {
              "protocolVersion": null
            }
            """);

    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    requestTree,
                    GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
                    WorkbookPlan.class,
                    GridGrindJsonProblemMessageSupport::invalidRequestPayload));

    ExplicitNullField problem = assertInstanceOf(ExplicitNullField.class, failure.requestProblem());
    assertEquals(
        "Field 'protocolVersion' must be omitted when absent; explicit null is not accepted.",
        failure.getMessage());
    assertEquals("protocolVersion", problem.jsonPathValue());
    assertEquals(Optional.of("protocolVersion"), failure.jsonPath());
  }

  @Test
  void decodeTreeDetectsMissingRequiredFieldsStructurally() throws IOException {
    JsonNode requestTree =
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER.readTree(
            """
            {
              "steps": []
            }
            """);

    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    requestTree,
                    GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
                    WorkbookPlan.class,
                    GridGrindJsonProblemMessageSupport::invalidRequestPayload));

    MissingRequiredField problem =
        assertInstanceOf(MissingRequiredField.class, failure.requestProblem());
    assertEquals("Missing required field 'protocolVersion'", failure.getMessage());
    assertEquals("protocolVersion", problem.jsonPathValue());
    assertEquals(Optional.of("protocolVersion"), failure.jsonPath());
  }
}
