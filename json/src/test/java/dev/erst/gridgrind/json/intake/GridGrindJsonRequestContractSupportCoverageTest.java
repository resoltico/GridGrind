package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.metadata.CatalogIgnored;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.plan.InspectionQuery;
import dev.erst.gridgrind.core.problem.MissingRequiredField;
import dev.erst.gridgrind.core.problem.MissingTypeDiscriminator;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.json.JsonMapper;

/** Checks canonical field presence before record binding. */
class GridGrindJsonRequestContractSupportCoverageTest {
  private record Shape(
      @CatalogIgnored String ignored, @WireDefault("true") boolean enabled, String required) {}

  @Test
  void requiredComponentsUseCanonicalNamesAndMetadataDefaults() {
    var empty = JsonMapper.builder().build().createObjectNode();
    assertEquals(
        "type",
        assertInstanceOf(
                MissingTypeDiscriminator.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        empty, InspectionQuery.class, "")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "required",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        empty, Shape.class, "")
                    .orElseThrow())
            .jsonPathValue());
    empty.put("required", "present");
    assertTrue(
        GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(empty, Shape.class, "")
            .isEmpty());
  }
}
