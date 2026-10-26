package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.text;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import org.junit.jupiter.api.Test;

/** Tests for CommentSpec record construction and engine conversion. */
class CommentSpecTest {
  @Test
  void hiddenCommentFactoryConvertsToEngineComment() {
    CommentSpec comment = CommentSpec.hidden(text("Review"), "GridGrind");

    assertFalse(comment.visible());
    assertEquals(
        new ResolvedComment("Review", "GridGrind", false),
        WorkbookCommandCellValueSpecConverter.toResolvedComment(comment));
  }

  @Test
  void validatesCommentInputs() {
    assertThrows(NullPointerException.class, () -> CommentSpec.plain(null, "GridGrind", true));
    assertThrows(
        IllegalArgumentException.class, () -> CommentSpec.plain(text(" "), "GridGrind", true));
    assertThrows(NullPointerException.class, () -> CommentSpec.plain(text("Review"), null, true));
    assertThrows(
        IllegalArgumentException.class, () -> CommentSpec.plain(text("Review"), " ", true));
  }
}
