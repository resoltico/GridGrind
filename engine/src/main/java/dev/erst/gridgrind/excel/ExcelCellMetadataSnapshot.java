package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import java.util.Objects;
import java.util.Optional;

/** Immutable optional hyperlink and comment facts captured for one analyzed cell. */
public record ExcelCellMetadataSnapshot(
    Optional<HyperlinkTarget> hyperlink, Optional<CommentReport> comment) {
  public ExcelCellMetadataSnapshot {
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");
    Objects.requireNonNull(comment, "comment must not be null");
  }

  /** Returns an empty cell-metadata snapshot with no hyperlink and no comment. */
  public static ExcelCellMetadataSnapshot empty() {
    return new ExcelCellMetadataSnapshot(Optional.empty(), Optional.empty());
  }

  /** Creates a metadata snapshot from optional hyperlink and comment values. */
  public static ExcelCellMetadataSnapshot of(
      Optional<HyperlinkTarget> hyperlink, Optional<CommentReport> comment) {
    return new ExcelCellMetadataSnapshot(hyperlink, comment);
  }
}
