package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ExplicitNullField;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.MissingRequiredField;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Direct coverage for request-shape detection at the tree-to-record decode seam. */
class GridGrindJsonCodecSupportTest {
  @Test
  void decodeTreeReturnsAValidBoundValue() throws IOException {
    JsonNode requestTree =
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER.readTree(
            """
            {
              "protocolVersion": 1,
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """);

    WorkbookPlan plan =
        GridGrindJsonCodecSupport.decodeTree(
            requestTree,
            GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
            WorkbookPlan.class,
            JsonReadFailures::syntax);

    assertEquals(0, plan.steps().size());
  }

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
                    JsonReadFailures::syntax));

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
                    JsonReadFailures::syntax));

    MissingRequiredField problem =
        assertInstanceOf(MissingRequiredField.class, failure.requestProblem());
    assertEquals("Missing required field 'protocolVersion'", failure.getMessage());
    assertEquals("protocolVersion", problem.jsonPathValue());
    assertEquals(Optional.of("protocolVersion"), failure.jsonPath());
  }

  @Test
  void codecSupportMapsStringGrammarFailuresAndNestedArrayNullsBeforeBinding() throws IOException {
    JsonMapper mapper = GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER;

    assertEquals(
        "{}",
        GridGrindJsonCodecSupport.readTree("{}", mapper, JsonReadFailures::syntax).toString());

    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJsonCodecSupport.readTree("{", mapper, JsonReadFailures::syntax)));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    mapper.readTree("[null]"), mapper, Object.class, JsonReadFailures::syntax)));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    mapper.nullNode(), mapper, Object.class, JsonReadFailures::syntax)));
  }

  @Test
  void rejectsNonObjectTopLevelPayloadsForWorkbookPlans() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.readValue(
                    "[]".getBytes(StandardCharsets.UTF_8),
                    GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
                    WorkbookPlan.class,
                    JsonReadFailures::syntax));
    assertEquals("JSON value has the wrong shape for this field", exception.getMessage());
  }
}
