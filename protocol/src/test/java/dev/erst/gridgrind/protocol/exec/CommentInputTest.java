package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.protocol.dto.CommentInput;
import org.junit.jupiter.api.Test;

/** Tests for CommentInput record construction and engine conversion. */
class CommentInputTest {
  @Test
  void defaultsVisibilityAndConvertsToEngineComment() {
    CommentInput comment = new CommentInput("Review", "GridGrind", null);

    assertFalse(comment.visible());
    assertEquals(
        new ExcelComment("Review", "GridGrind", false),
        WorkbookCommandConverter.toExcelComment(comment));
  }

  @Test
  void validatesCommentInputs() {
    assertThrows(NullPointerException.class, () -> new CommentInput(null, "GridGrind", true));
    assertThrows(IllegalArgumentException.class, () -> new CommentInput(" ", "GridGrind", true));
    assertThrows(NullPointerException.class, () -> new CommentInput("Review", null, true));
    assertThrows(IllegalArgumentException.class, () -> new CommentInput("Review", " ", true));
  }
}
