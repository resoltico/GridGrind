package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;

/** Centralized Excel cell-text ceiling checks for authored and reconstructed sheet content. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelCellTextLimits {
  public static final int MAX_CELL_TEXT_LENGTH =
      SpreadsheetVersion.EXCEL2007.getMaxTextLength(); // LIM-010

  private ExcelCellTextLimits() {}

  public static void requireSupportedLength(String text, String fieldName) {
    Objects.requireNonNull(text, fieldName + " must not be null");
    if (text.length() > MAX_CELL_TEXT_LENGTH) {
      throw new IllegalArgumentException(
          fieldName
              + " must not exceed "
              + MAX_CELL_TEXT_LENGTH
              + " characters (Excel cell text limit)");
    }
  }
}
