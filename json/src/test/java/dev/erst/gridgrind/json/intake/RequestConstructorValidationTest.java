package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Checks canonical constructors independently against the final typed codec and owned failures. */
class RequestConstructorValidationTest {
  @Test
  void constructorPrevalidationRejectsDomainValuesBeforeAnyJacksonBinding() {
    var node = new RequestJsonString(0, "bad id");
    List<RequestBindingFailure> failures = new ArrayList<>();
    var result =
        RequestConstructorValidation.validate(
            node, dev.erst.gridgrind.core.vocabulary.StepId.class, "steps[0].stepId", failures);
    assertTrue(result.isEmpty());
    assertEquals(1, failures.size());
    assertEquals("steps[0].stepId", failures.getFirst().jsonPath());
    assertInstanceOf(
        dev.erst.gridgrind.core.problem.InvalidRequestException.class,
        failures.getFirst().exception());
  }

  @Test
  void everyExampleMatchesIndependentFragmentCodecsWithoutAnotherRootDecoder() throws IOException {
    Path root = Path.of(System.getProperty("gridgrind.root.dir"));
    try (var paths = Files.list(root.resolve("examples"))) {
      List<Path> examples =
          paths.filter(path -> path.toString().endsWith(".json")).sorted().toList();
      assertEquals(14, examples.size());
      List<RequestBindingFailure> problems = new ArrayList<>();
      var stepListType =
          new tools.jackson.core.type.TypeReference<
              List<dev.erst.gridgrind.core.plan.WorkbookStep>>() {};
      for (Path file : examples) {
        problems.clear();
        byte[] bytes = Files.readAllBytes(file);
        var parsed = TolerantRequestJsonParser.parse(bytes);
        assertTrue(parsed.problems().isEmpty());
        var actual =
            RequestConstructorValidation.validate(parsed.root(), WorkbookPlan.class, "", problems);
        assertTrue(problems.isEmpty(), () -> file + ": " + problems);
        var tree = RequestFragmentBinder.toJsonNode(parsed.root());
        var mapper = dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.JSON_MAPPER;
        var plan = actual.orElseThrow();
        assertEquals(
            mapper.treeToValue(
                tree.path("protocolVersion"),
                dev.erst.gridgrind.core.vocabulary.ProtocolVersion.class),
            plan.protocolVersion());
        assertEquals(
            mapper.treeToValue(tree.path("source"), WorkbookPlan.WorkbookSource.class),
            plan.source());
        assertEquals(
            mapper.treeToValue(tree.path("persistence"), WorkbookPlan.WorkbookPersistence.class),
            plan.persistence());
        if (tree.has("execution")) {
          assertEquals(
              mapper.treeToValue(
                  tree.path("execution"), dev.erst.gridgrind.core.plan.ExecutionPolicySpec.class),
              plan.execution());
        } else {
          assertTrue(plan.execution().isDefault());
        }
        if (tree.has("formulaEnvironment")) {
          assertEquals(
              mapper.treeToValue(
                  tree.path("formulaEnvironment"),
                  dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec.class),
              plan.formulaEnvironment());
        } else {
          assertTrue(plan.formulaEnvironment().isEmpty());
        }
        assertEquals(
            mapper.readValue(mapper.writeValueAsBytes(tree.path("steps")), stepListType),
            plan.steps());
      }
    }
  }
}
