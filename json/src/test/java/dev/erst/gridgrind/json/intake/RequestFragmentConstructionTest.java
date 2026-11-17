package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestExplicitNullField;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;

/** Checks single canonical construction and metadata-owned shape and omission defaults. */
class RequestFragmentConstructionTest {
  private static final AtomicInteger CONSTRUCTIONS = new AtomicInteger();

  @Test
  void bindsTheCanonicalConstructorResultOnce() {
    CONSTRUCTIONS.set(0);
    var node =
        new RequestJsonObject(
            0, List.of(new RequestJsonMember("value", 1, new RequestJsonString(9, "kept"))));
    List<RequestBindingFailure> failures = new ArrayList<>();
    var bound = RequestFragmentBinder.bindNode(node, Counted.class, "field", List.of(), failures);
    assertTrue(failures.isEmpty());
    assertEquals("kept", bound.orElseThrow().value());
    assertEquals(1, CONSTRUCTIONS.get());
  }

  @Test
  void shapeChecksUseDeclaredFieldsWithoutAnExportedSchemaDefinition() {
    var node =
        new RequestJsonObject(
            0, List.of(new RequestJsonMember("count", 1, new RequestJsonNumber(9, "7"))));
    List<RequestStructuralProblem> problems = new ArrayList<>();
    RequestSchemaStructure.validateRecord(node, IndependentShape.class, "field", 0, problems);
    assertTrue(problems.isEmpty(), problems.toString());
    List<RequestBindingFailure> failures = new ArrayList<>();
    var bound =
        RequestFragmentBinder.bindNode(node, IndependentShape.class, "field", problems, failures);
    assertEquals(new IndependentShape(7, false), bound.orElseThrow());
    var explicitNull =
        new RequestJsonObject(
            0,
            List.of(
                new RequestJsonMember("count", 1, new RequestJsonNumber(9, "7")),
                new RequestJsonMember("enabled", 12, new RequestJsonNull(22))));
    RequestSchemaStructure.validateRecord(
        explicitNull, IndependentShape.class, "field", 0, problems);
    assertEquals(1, problems.size());
    assertInstanceOf(RequestExplicitNullField.class, problems.getFirst());
    assertTrue(
        RequestFragmentBinder.bindNode(
                explicitNull, IndependentShape.class, "field", problems, failures)
            .isEmpty());
  }

  @SuppressWarnings("PMD.PublicMemberInNonPublicType")
  // Canonical fragment binding selects a public constructor; this test fixture must retain that
  // access.
  public record Counted(String value) {
    public Counted {
      CONSTRUCTIONS.incrementAndGet();
    }
  }

  public record IndependentShape(int count, @WireDefault(value = "false") boolean enabled) {}
}
