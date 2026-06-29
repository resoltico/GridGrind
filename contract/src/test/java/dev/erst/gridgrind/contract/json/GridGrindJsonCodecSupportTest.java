package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

/** Direct coverage for request-shape normalization at the tree-to-record decode seam. */
class GridGrindJsonCodecSupportTest {
  @Test
  void preservesExplicitNullFailuresWhenTheAuthoredTreeActuallyContainsThatPath()
      throws IOException {
    JsonNode requestTree =
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER.readTree(
            """
            {
              "protocolVersion": "V1"
            }
            """);
    InvalidRequestShapeException explicitNullFailure =
        new InvalidRequestShapeException(
            GridGrindRequestProblemSupport.explicitNullFieldMessage("protocolVersion"),
            Optional.of("protocolVersion"),
            Optional.empty(),
            Optional.empty(),
            null);

    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    requestTree,
                    GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
                    WorkbookPlan.class,
                    exception -> explicitNullFailure));

    assertSame(explicitNullFailure, failure);
  }

  @Test
  void rewritesExplicitNullFailuresToMissingRequiredWhenOnlyTheMessageCarriesTheLeafField()
      throws IOException {
    JsonNode requestTree =
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER.readTree(
            """
            {
              "steps": []
            }
            """);
    InvalidRequestShapeException explicitNullFailure =
        new InvalidRequestShapeException(
            GridGrindRequestProblemSupport.explicitNullFieldMessage("protocolVersion"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            null);

    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.decodeTree(
                    requestTree,
                    GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
                    WorkbookPlan.class,
                    exception -> explicitNullFailure));

    org.junit.jupiter.api.Assertions.assertEquals(
        "Missing required field 'protocolVersion'", failure.getMessage());
    org.junit.jupiter.api.Assertions.assertEquals(
        Optional.of("protocolVersion"), failure.jsonPath());
  }
}
