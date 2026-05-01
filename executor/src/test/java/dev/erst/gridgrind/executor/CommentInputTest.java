package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.text;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.excel.ExcelComment;
import org.junit.jupiter.api.Test;

/** Tests for CommentInput record construction and engine conversion. */
class CommentInputTest {
  @Test
  void hiddenCommentFactoryConvertsToEngineComment() {
    CommentInput comment = CommentInput.hidden(text("Review"), "GridGrind");

    assertFalse(comment.visible());
    assertEquals(
        new ExcelComment("Review", "GridGrind", false),
        WorkbookCommandConverter.toExcelComment(comment));
  }

  @Test
  void validatesCommentInputs() {
    assertThrows(NullPointerException.class, () -> CommentInput.plain(null, "GridGrind", true));
    assertThrows(
        IllegalArgumentException.class, () -> CommentInput.plain(text(" "), "GridGrind", true));
    assertThrows(NullPointerException.class, () -> CommentInput.plain(text("Review"), null, true));
    assertThrows(
        IllegalArgumentException.class, () -> CommentInput.plain(text("Review"), " ", true));
  }
}
