package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/** One ignored-error block anchored to one A1-style range plus one or more error families. */
public record IgnoredErrorInput(RangeRef range, List<ExcelIgnoredErrorType> errorTypes) {
  public IgnoredErrorInput {
    Objects.requireNonNull(range, "range must not be null");
    errorTypes = List.copyOf(Objects.requireNonNull(errorTypes, "errorTypes must not be null"));
    if (errorTypes.isEmpty()) {
      throw new IllegalArgumentException("errorTypes must not be empty");
    }
    if (errorTypes.size() != new LinkedHashSet<>(errorTypes).size()) {
      throw new IllegalArgumentException("errorTypes must not contain duplicates");
    }
  }
}
