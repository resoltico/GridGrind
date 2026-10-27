package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.layout.WindowDimensionsReport;
import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.problem.HostContext;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;
import tools.jackson.databind.node.ObjectNode;

/** Verifies current tag spelling and rejects the removed discriminator names. */
class CanonicalUnionTagsTest {
  @Test
  void windowsUseTypeAndRejectShape() {
    verify(
        new WindowReport.Sparse(
            new SheetLabel("Data"),
            new CellAddress("A1"),
            new WindowDimensionsReport(1, 1),
            List.of()),
        WindowReport.class,
        "type",
        "shape",
        "SPARSE");
  }

  @Test
  void contextsUseTypeAndRejectStage() {
    verify(new HostContext("cli"), ProblemContext.class, "type", "stage", "HOST_RUNTIME");
  }

  @Test
  void assertionsUseStatusAndRejectOutcome() {
    verify(
        new AssertionResult.Passed(new StepId("check"), "EXPECT_CELL_VALUE"),
        AssertionResult.class,
        "status",
        "outcome",
        "PASSED");
  }

  private static <T> void verify(T value, Class<T> union, String tag, String removed, String id) {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    byte[] valid = mapper.writeValueAsBytes(value);
    ObjectNode node = (ObjectNode) mapper.readTree(valid);
    assertEquals(id, node.path(tag).asText());
    assertFalse(node.has(removed));
    assertEquals(value, mapper.readValue(valid, union));
    node.remove(tag);
    node.put(removed, id);
    assertThrows(
        JacksonException.class, () -> mapper.readValue(mapper.writeValueAsBytes(node), union));
  }
}
