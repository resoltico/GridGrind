package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.NamedRangeBackingKind;
import java.util.Objects;

/** One named-range surface entry classified by scope and backing kind. */
public record NamedRangeSurfaceEntryReport(
    String name, NamedRangeScopeReport scope, String refersToFormula, NamedRangeBackingKind kind) {
  public NamedRangeSurfaceEntryReport {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(scope, "scope must not be null");
    Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
    Objects.requireNonNull(kind, "kind must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
  }
}
