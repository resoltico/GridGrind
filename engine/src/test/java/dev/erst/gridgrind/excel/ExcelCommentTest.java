package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelComment record construction. */
class ExcelCommentTest {
  @Test
  void validatesCommentInputs() {
    assertDoesNotThrow(() -> new ExcelComment("Review", "GridGrind", true));
    assertThrows(NullPointerException.class, () -> new ExcelComment(null, "GridGrind", false));
    assertThrows(IllegalArgumentException.class, () -> new ExcelComment(" ", "GridGrind", false));
    assertThrows(NullPointerException.class, () -> new ExcelComment("Review", null, false));
    assertThrows(IllegalArgumentException.class, () -> new ExcelComment("Review", " ", false));
  }
}
