package dev.erst.gridgrind.excel;

import java.util.Optional;

/** Immutable optional hyperlink and comment facts captured for one analyzed cell. */
public record ExcelCellMetadataSnapshot(
    Optional<ExcelHyperlink> hyperlink, Optional<ExcelComment> comment) {
  public ExcelCellMetadataSnapshot {
    hyperlink = hyperlink == null ? Optional.empty() : hyperlink;
    comment = comment == null ? Optional.empty() : comment;
  }

  /** Returns an empty cell-metadata snapshot with no hyperlink and no comment. */
  public static ExcelCellMetadataSnapshot empty() {
    return new ExcelCellMetadataSnapshot(Optional.empty(), Optional.empty());
  }

  /** Creates a metadata snapshot from possibly-null hyperlink and comment values. */
  public static ExcelCellMetadataSnapshot of(ExcelHyperlink hyperlink, ExcelComment comment) {
    return new ExcelCellMetadataSnapshot(
        Optional.ofNullable(hyperlink), Optional.ofNullable(comment));
  }
}
