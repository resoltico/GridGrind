package dev.erst.gridgrind.core.model.drawing;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ResolvedComment record construction. */
class ResolvedCommentTest {
  @Test
  void validatesCommentInputs() {
    assertDoesNotThrow(() -> new ResolvedComment("Review", "GridGrind", true));
    assertThrows(NullPointerException.class, () -> new ResolvedComment(null, "GridGrind", false));
    assertThrows(
        IllegalArgumentException.class, () -> new ResolvedComment(" ", "GridGrind", false));
    assertThrows(NullPointerException.class, () -> new ResolvedComment("Review", null, false));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedComment("Review", " ", false));
  }
}
