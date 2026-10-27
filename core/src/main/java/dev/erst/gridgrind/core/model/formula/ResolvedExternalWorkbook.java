package dev.erst.gridgrind.core.model.formula;

import java.nio.file.Path;
import java.util.Objects;

/** One authorized, materialized external workbook available to formula evaluation. */
public record ResolvedExternalWorkbook(String workbookName, Path path) {
  public ResolvedExternalWorkbook {
    Objects.requireNonNull(workbookName, "workbookName must not be null");
    if (workbookName.isBlank())
      throw new IllegalArgumentException("workbookName must not be blank");
    Objects.requireNonNull(path, "path must not be null");
  }
}
