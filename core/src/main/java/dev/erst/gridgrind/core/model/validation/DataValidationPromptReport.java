package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;

/** Observed prompt-box text and display state, without authored source references. */
@CatalogType(
    id = "DataValidationPromptReport",
    summary = "Observed prompt-box text and display state.",
    displayOrder = 108)
public record DataValidationPromptReport(String title, String text, boolean showPromptBox) {
  public DataValidationPromptReport {
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(text, "text must not be null");
  }
}
