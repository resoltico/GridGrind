package dev.erst.gridgrind.contract.catalog;

/** Public metadata for one generated built-in example workbook plan. */
public record ShippedExampleEntry(String id, String fileName, String summary) {
  public ShippedExampleEntry {
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    fileName = CatalogRecordValidation.requireNonBlank(fileName, "fileName");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    if (!fileName.endsWith(".json")) {
      throw new IllegalArgumentException("fileName must end with .json");
    }
  }
}
