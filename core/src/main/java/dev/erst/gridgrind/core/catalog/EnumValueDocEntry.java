package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Machine-readable summary for one published enum token on a catalog field. */
public record EnumValueDocEntry(String value, String summary) {
  public EnumValueDocEntry {
    value = TextValidation.requireNonBlank(value, "value");
    summary = TextValidation.requireNonBlank(summary, "summary");
  }
}
