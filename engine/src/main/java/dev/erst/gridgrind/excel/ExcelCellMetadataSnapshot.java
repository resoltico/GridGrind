package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Immutable optional hyperlink and comment facts captured for one analyzed cell. */
public record ExcelCellMetadataSnapshot(
    Optional<ExcelHyperlink> hyperlink, Optional<ExcelCommentSnapshot> comment) {
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
      Optional<ExcelHyperlink> hyperlink, Optional<ExcelCommentSnapshot> comment) {
    return new ExcelCellMetadataSnapshot(hyperlink, comment);
  }
}
