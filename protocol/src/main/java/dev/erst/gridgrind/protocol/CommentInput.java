package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelComment;
import java.util.Objects;

/** Protocol-facing plain-text comment payload used by cell comment operations. */
public record CommentInput(String text, String author, Boolean visible) {
  public CommentInput {
    text = requireNonBlank(text, "text");
    author = requireNonBlank(author, "author");
    visible = visible == null ? Boolean.FALSE : visible;
  }

  /** Converts this transport comment into the workbook-core representation. */
  public ExcelComment toExcelComment() {
    return new ExcelComment(text, author, visible);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
