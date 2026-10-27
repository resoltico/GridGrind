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
  void everyExampleConstructsTheSameCanonicalPlanBeforeJacksonBinding() throws IOException {
    Path root = Path.of(System.getProperty("gridgrind.root.dir"));
    try (var paths = Files.list(root.resolve("examples"))) {
      List<Path> examples =
          paths.filter(path -> path.toString().endsWith(".json")).sorted().toList();
      assertEquals(14, examples.size());
      List<RequestBindingFailure> problems = new ArrayList<>();
      for (Path file : examples) {
        problems.clear();
        byte[] bytes = Files.readAllBytes(file);
        var parsed = TolerantRequestJsonParser.parse(bytes);
        assertTrue(parsed.problems().isEmpty());
        var actual =
            RequestConstructorValidation.validate(parsed.root(), WorkbookPlan.class, "", problems);
        assertTrue(problems.isEmpty(), () -> file + ": " + problems);
        assertEquals(
            dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER
                .treeToValue(RequestFragmentBinder.toJsonNode(parsed.root()), WorkbookPlan.class),
            actual.orElseThrow(),
            file.toString());
      }
    }
  }
}
