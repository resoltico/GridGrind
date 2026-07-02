package dev.erst.gridgrind.contract.catalog;

/** Machine-readable summary for one published enum token on a catalog field. */
public record EnumValueDocEntry(String value, String summary) {
  public EnumValueDocEntry {
    value = CatalogRecordValidation.requireNonBlank(value, "value");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
  }
}
