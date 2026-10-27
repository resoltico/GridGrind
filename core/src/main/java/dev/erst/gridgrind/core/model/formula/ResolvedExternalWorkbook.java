package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.nio.file.Path;
import java.util.Objects;

/** One authorized, materialized external workbook available to formula evaluation. */
public record ResolvedExternalWorkbook(String workbookName, Path path) {
  public ResolvedExternalWorkbook {
    workbookName = TextValidation.requireNonBlank(workbookName, "workbookName");
    Objects.requireNonNull(path, "path must not be null");
  }
}
