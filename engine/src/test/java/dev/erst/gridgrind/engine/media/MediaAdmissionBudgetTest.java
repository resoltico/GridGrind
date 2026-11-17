package dev.erst.gridgrind.engine.media;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import org.junit.jupiter.api.Test;

/**
 * Shared counters consume actual occurrence/work values without allocating represented raster data.
 */
class MediaAdmissionBudgetTest {
  @Test
  void occurrenceAndRunBytesAreDistinctAndNeverRefundFailure() {
    var oversized = new MediaAdmissionBudget();
    assertLimit(Limit.ID_AUTHORED_MEDIA_BYTES, () -> oversized.reserveOccurrence(33_554_433));
    var run = new MediaAdmissionBudget();
    run.reserveOccurrence(33_554_432);
    run.reserveOccurrence(33_554_432);
    assertEquals(67_108_864, run.mediaUseBytes());
    assertLimit(Limit.ID_AUTHORED_MEDIA_TOTAL_BYTES, () -> run.reserveOccurrence(1));
    assertLimit(Limit.ID_AUTHORED_MEDIA_TOTAL_BYTES, run::payloadAllowance);
  }

  @Test
  void rasterExpansionAndInspectionChargeAggregatesAndCheckedProducts() {
    var pixels = new MediaAdmissionBudget();
    pixels.raster(16_777_216, 1);
    assertLimit(Limit.ID_MEDIA_RASTER_PIXELS, () -> pixels.raster(1, 1));
    assertLimit(
        Limit.ID_MEDIA_RASTER_PIXELS, () -> new MediaAdmissionBudget().raster(Long.MAX_VALUE, 2));
    var expanded = new MediaAdmissionBudget();
    expanded.expand(134_217_728);
    assertLimit(Limit.ID_MEDIA_EXPANDED_BYTES, () -> expanded.expand(1));
    var inspected = new MediaAdmissionBudget();
    inspected.inspect(268_435_456);
    assertLimit(Limit.ID_MEDIA_INSPECTION_BYTES, () -> inspected.inspect(1));
  }

  @Test
  void structuralAndLexicalBoundsIncludeTheExactLimit() {
    var records = new MediaAdmissionBudget();
    for (int i = 0; i < 65_536; i++) records.record();
    assertLimit(Limit.ID_MEDIA_STRUCTURE_RECORDS, records::record);
    var tokens = new MediaAdmissionBudget();
    for (int i = 0; i < 1_000_000; i++) tokens.token();
    assertLimit(Limit.ID_EPS_SYNTAX_TOKENS, tokens::token);
    var depth = new MediaAdmissionBudget();
    assertDoesNotThrow(() -> depth.depth(8));
    assertLimit(Limit.ID_MEDIA_CONTAINER_DEPTH, () -> depth.depth(9));
    var syntax = new MediaAdmissionBudget();
    assertDoesNotThrow(() -> syntax.syntaxDepth(256));
    assertLimit(Limit.ID_EPS_SYNTAX_DEPTH, () -> syntax.syntaxDepth(257));
  }

  private static void assertLimit(String id, org.junit.jupiter.api.function.Executable call) {
    var refusal = assertThrows(InvalidRequestException.class, call);
    var field = (FieldValidationProblem) refusal.requestProblem();
    assertEquals(FieldValidationMediaRule.RESOURCE_LIMIT, field.rule());
    assertEquals(id, field.operands().getFirst());
  }
}
