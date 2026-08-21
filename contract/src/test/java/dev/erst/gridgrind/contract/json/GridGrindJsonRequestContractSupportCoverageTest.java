package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Covers annotation-carrier and discriminator branches for request-contract shape detection. */
class GridGrindJsonRequestContractSupportCoverageTest {
  @Test
  void missingRequiredComponentDetectionHonorsDiscriminatorsAndRecordAnnotationCarriers() {
    JsonNode emptyObject = JsonMapper.builder().build().createObjectNode();

    assertEquals(
        "type",
        assertInstanceOf(
                MissingTypeDiscriminator.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, InspectionQuery.class, "")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "visible",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, IgnoredComponentRecord.class, "")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "visible",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, IgnoredAccessorRecord.class, "")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "componentWire",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, ComponentPropertyRecord.class, "")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "componentWire",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, ComponentPropertyRecord.class, "componentWire")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "required",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, BlankComponentPropertyRecord.class, "")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "accessorWire",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, AccessorPropertyRecord.class, "")
                    .orElseThrow())
            .jsonPathValue());
    assertEquals(
        "required",
        assertInstanceOf(
                MissingRequiredField.class,
                GridGrindJsonRequestContractSupport.missingRequiredComponentProblem(
                        emptyObject, BlankAccessorPropertyRecord.class, "")
                    .orElseThrow())
            .jsonPathValue());
  }

  private record IgnoredComponentRecord(@JsonIgnore String ignored, String visible) {}

  private record IgnoredAccessorRecord(String ignored, String visible) {
    @JsonIgnore
    @SuppressWarnings("UnusedMethod")
    @Override
    public String ignored() {
      return ignored;
    }
  }

  private record ComponentPropertyRecord(@JsonProperty("componentWire") String required) {}

  private record BlankComponentPropertyRecord(@JsonProperty("") String required) {}

  private record AccessorPropertyRecord(String required) {
    @JsonProperty("accessorWire")
    @Override
    public String required() {
      return required;
    }
  }

  private record BlankAccessorPropertyRecord(String required) {
    @JsonProperty("")
    @Override
    public String required() {
      return required;
    }
  }
}
