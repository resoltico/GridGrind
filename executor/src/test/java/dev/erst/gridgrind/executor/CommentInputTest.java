package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.text;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.excel.ExcelComment;
import org.junit.jupiter.api.Test;

/** Tests for CommentInput record construction and engine conversion. */
class CommentInputTest {
  @Test
  void defaultsVisibilityAndConvertsToEngineComment() {
    CommentInput comment = new CommentInput(text("Review"), "GridGrind", null);

    assertFalse(comment.visible());
    assertEquals(
        new ExcelComment("Review", "GridGrind", false),
        WorkbookCommandConverter.toExcelComment(comment));
  }

  @Test
  void validatesCommentInputs() {
    assertThrows(NullPointerException.class, () -> new CommentInput(null, "GridGrind", true));
    assertThrows(
        IllegalArgumentException.class, () -> new CommentInput(text(" "), "GridGrind", true));
    assertThrows(NullPointerException.class, () -> new CommentInput(text("Review"), null, true));
    assertThrows(IllegalArgumentException.class, () -> new CommentInput(text("Review"), " ", true));
  }
}
