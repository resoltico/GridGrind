package dev.erst.gridgrind.excel;

/** Centralized worksheet hyperlink-capacity checks before new hyperlinks are authored. */
final class ExcelHyperlinkLimits {
  static final int MAX_HYPERLINKS_PER_SHEET = 65530; // LIM-012

  private ExcelHyperlinkLimits() {}

  static void requireWorksheetHyperlinkCapacity(int existingHyperlinkCount) {
    if (existingHyperlinkCount >= MAX_HYPERLINKS_PER_SHEET) {
      throw new IllegalArgumentException(
          "sheet cannot contain more than "
              + MAX_HYPERLINKS_PER_SHEET
              + " hyperlinks (Excel worksheet hyperlink limit)");
    }
  }
}
