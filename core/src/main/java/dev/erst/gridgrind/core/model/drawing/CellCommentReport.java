package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import java.util.Objects;

/** Comment metadata associated with one concrete cell address. */
public record CellCommentReport(CellAddress address, CommentReport comment) {
  public CellCommentReport {
    Objects.requireNonNull(address, "address must not be null");
    Objects.requireNonNull(comment, "comment must not be null");
  }
}
