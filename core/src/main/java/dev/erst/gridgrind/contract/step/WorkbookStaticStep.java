package dev.erst.gridgrind.contract.step;

import java.util.Objects;
import java.util.Optional;

/** One independently bound authored step retained at its original request index. */
public record WorkbookStaticStep(int index, Optional<WorkbookStep> value) {
  public WorkbookStaticStep {
    if (index < 0) {
      throw new IllegalArgumentException("index must not be negative");
    }
    Objects.requireNonNull(value, "value must not be null");
    value.ifPresent(candidate -> Objects.requireNonNull(candidate, "value must not contain null"));
  }
}
