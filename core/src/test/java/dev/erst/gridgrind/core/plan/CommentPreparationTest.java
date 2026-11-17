package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import org.junit.jupiter.api.Test;

/** Tests for CommentSpec record construction and source-free preparation. */
class CommentPreparationTest {
  @Test
  void hiddenCommentFactoryPreparesResolvedComment() {
    CommentSpec comment =
        CommentSpec.hidden(
            dev.erst.gridgrind.core.source.TextSourceSpec.inline("Review"), "GridGrind");

    assertFalse(comment.visible());
    assertEquals(
        new ResolvedComment("Review", "GridGrind", false),
        MutationPreparation.toResolvedComment(comment));
  }

  @Test
  void validatesCommentInputs() {
    assertThrows(NullPointerException.class, () -> CommentSpec.plain(null, "GridGrind", true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CommentSpec.plain(
                dev.erst.gridgrind.core.source.TextSourceSpec.inline(" "), "GridGrind", true));
    assertThrows(
        NullPointerException.class,
        () ->
            CommentSpec.plain(
                dev.erst.gridgrind.core.source.TextSourceSpec.inline("Review"), null, true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CommentSpec.plain(
                dev.erst.gridgrind.core.source.TextSourceSpec.inline("Review"), " ", true));
  }
}
