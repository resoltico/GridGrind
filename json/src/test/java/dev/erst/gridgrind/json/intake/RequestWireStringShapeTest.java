package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.catalog.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.core.catalog.FieldShape;
import dev.erst.gridgrind.core.catalog.ScalarType;
import dev.erst.gridgrind.core.problem.RequestExplicitNullField;
import dev.erst.gridgrind.core.problem.RequestMalformedScalar;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.WireString;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Checks the same scalar contract through catalog metadata and tolerant request validation. */
class RequestWireStringShapeTest {
  @Test
  @SuppressWarnings(
      "PMD.AvoidInstantiatingObjectsInLoops") // Findings must be isolated for each type.
  void everyDomainStringHasAScalarCatalogShapeAndAStringTokenShape() {
    for (Class<?> type : WireString.class.getPermittedSubclasses()) {
      assertEquals(
          new FieldShape.Scalar(ScalarType.STRING), CatalogFieldMetadataSupport.fieldShape(type));
      List<RequestStructuralProblem> problems = new ArrayList<>();
      RequestNodeValidator.validateNode(
          new RequestJsonString(7, "value"), type, "field", 7, problems);
      assertTrue(problems.isEmpty(), type.getName());
      RequestNodeValidator.validateNode(
          new RequestJsonBoolean(8, true), type, "field", 8, problems);
      assertEquals(List.of(new RequestMalformedScalar("field", "a JSON string", 8)), problems);
    }
  }

  @Test
  void explicitNullRemainsAnIntakeProblemEvenForOptionalDomainValues() {
    Type optionalSheet = OptionalSheet.class.getRecordComponents()[0].getGenericType();
    List<RequestStructuralProblem> problems = new ArrayList<>();
    RequestNodeValidator.validateNode(
        new RequestJsonNull(12), optionalSheet, "sheet", 12, problems);
    assertEquals(List.of(new RequestExplicitNullField("sheet", 12)), problems);
  }

  @Test
  void planIdentityBindsAsAScalarAndMayBeOmitted() throws java.io.IOException {
    var absent = dev.erst.gridgrind.json.codec.GridGrindJson.readRequest(planDocument(""));
    assertTrue(absent.planId().isEmpty());
    var present =
        dev.erst.gridgrind.json.codec.GridGrindJson.readRequest(
            planDocument(",\"planId\":\"plan 1\""));
    assertEquals("plan 1", present.planId().orElseThrow().value());
    byte[] bytes = dev.erst.gridgrind.json.codec.GridGrindJsonOutput.writeRequestBytes(present);
    assertEquals(
        "plan 1",
        dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER
            .readTree(bytes)
            .path("planId")
            .stringValue());
    assertEquals(present, dev.erst.gridgrind.json.codec.GridGrindJson.readRequest(bytes));
  }

  @Test
  void invalidPlanIdentityRetainsItsAuthoredPathAndTokenOffset() {
    var explicitNull =
        dev.erst.gridgrind.json.codec.GridGrindJson.analyzeRequest(
            planDocument(",\"planId\":null").getBytes(java.nio.charset.StandardCharsets.UTF_8));
    assertFalse(explicitNull.isBindable());
    assertEquals(1, explicitNull.structuralProblems().size());
    assertEquals(Optional.of("planId"), explicitNull.structuralProblems().getFirst().jsonPath());
    assertTrue(explicitNull.structuralProblems().getFirst().byteOffset().isPresent());
    var blank =
        dev.erst.gridgrind.json.codec.GridGrindJson.analyzeRequest(
            planDocument(",\"planId\":\" \"").getBytes(java.nio.charset.StandardCharsets.UTF_8));
    assertFalse(blank.isBindable());
    assertEquals(1, blank.bindingFailures().size());
    assertEquals("planId", blank.bindingFailures().getFirst().jsonPath());
    assertTrue(blank.bindingFailures().getFirst().byteOffset().isPresent());
    assertEquals(
        "planId must not be blank", blank.bindingFailures().getFirst().exception().getMessage());
  }

  @Test
  void stepIdentityBindsAndRendersAsTheAuthoredScalar() throws java.io.IOException {
    String request = stepDocument("\"read-1\"");
    var plan = dev.erst.gridgrind.json.codec.GridGrindJson.readRequest(request);
    assertEquals("read-1", plan.steps().getFirst().stepId().value());
    byte[] bytes = dev.erst.gridgrind.json.codec.GridGrindJsonOutput.writeRequestBytes(plan);
    assertEquals(
        "read-1",
        dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER
            .readTree(bytes)
            .path("steps")
            .path(0)
            .path("stepId")
            .stringValue());
    assertEquals(plan, dev.erst.gridgrind.json.codec.GridGrindJson.readRequest(bytes));
  }

  @Test
  void invalidStepIdentityRetainsTheEnvelopeMemberLocation() {
    var intake =
        dev.erst.gridgrind.json.codec.GridGrindJson.analyzeRequest(
            stepDocument("\"bad id\"").getBytes(java.nio.charset.StandardCharsets.UTF_8));
    assertFalse(intake.isBindable());
    assertEquals(1, intake.bindingFailures().size());
    assertEquals("steps[0].stepId", intake.bindingFailures().getFirst().jsonPath());
    assertTrue(intake.bindingFailures().getFirst().byteOffset().isPresent());
    assertEquals(
        "stepId must match [A-Za-z0-9._-]+",
        intake.bindingFailures().getFirst().exception().getMessage());
  }

  private static String stepDocument(String stepId) {
    return """
        {"protocolVersion":"V3","source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[{"stepId":%s,"target":{"type":"WORKBOOK_CURRENT"},"query":{"type":"GET_WORKBOOK_SUMMARY"}}]}
        """
        .formatted(stepId);
  }

  private static String planDocument(String extra) {
    return """
        {"protocolVersion":"V3","source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":[]%s}
        """
        .formatted(extra);
  }

  private record OptionalSheet(Optional<SheetName> sheet) {}
}
