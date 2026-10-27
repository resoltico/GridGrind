package dev.erst.gridgrind.core.limits;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import java.util.HashSet;
import org.junit.jupiter.api.Test;

/** Checks stable catalog identity and machine-readable problem references. */
class LimitTest {
  @Test
  void everyDocumentedIdentityHasOneCompleteEntry() {
    assertEquals(41, Limit.all().size());
    var ids = new HashSet<String>();
    for (Limit limit : Limit.all()) {
      assertTrue(ids.add(limit.id()), limit.id());
      assertTrue(limit.id().matches("LIM-[0-9]{3}[A-Z]?"), limit.id());
      assertTrue(!limit.title().isBlank() && !limit.value().isBlank(), limit.id());
      for (String code : limit.problemCodes()) GridGrindProblemCode.valueOf(code);
      assertEquals(limit, Limit.find(limit.id()).orElseThrow());
    }
    assertTrue(Limit.find("unknown").isEmpty());
  }

  @Test
  void numericPublicationUsesTheEnforcementValues() {
    assertEquals(Limit.MAX_READ_CELLS + " cells", Limit.CELL_RETURNING_READ_COUNT.value());
    assertEquals((Limit.MAX_ROW_INDEX + 1) + " rows", Limit.MAXIMUM_ROWS_PER_WORKSHEET.value());
    assertEquals(
        (Limit.MAX_COLUMN_INDEX + 1) + " columns", Limit.MAXIMUM_COLUMNS_PER_WORKSHEET.value());
    assertEquals("LIM-006A", Limit.STEP_ID_TOKEN_SHAPE.id());
  }

  @Test
  void workbookCoordinateAndStyleCeilingsMatchTheEstablishedXlsxContract() {
    assertEquals(1_048_575, Limit.MAX_ROW_INDEX);
    assertEquals(16_383, Limit.MAX_COLUMN_INDEX);
    assertEquals(64_000, Limit.MAX_CELL_STYLES);
  }
}
