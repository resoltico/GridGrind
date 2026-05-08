package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** One inferred schema column with header text and observed value-type counts. */
public record SchemaColumnReport(
    int columnIndex,
    String columnAddress,
    String headerDisplayValue,
    int populatedCellCount,
    int blankCellCount,
    List<TypeCountReport> observedTypes,
    @Nullable String dominantType) {
  public SchemaColumnReport {
    if (columnIndex < 0) {
      throw new IllegalArgumentException("columnIndex must not be negative");
    }
    Objects.requireNonNull(columnAddress, "columnAddress must not be null");
    Objects.requireNonNull(headerDisplayValue, "headerDisplayValue must not be null");
    if (columnAddress.isBlank()) {
      throw new IllegalArgumentException("columnAddress must not be blank");
    }
    if (populatedCellCount < 0) {
      throw new IllegalArgumentException("populatedCellCount must not be negative");
    }
    if (blankCellCount < 0) {
      throw new IllegalArgumentException("blankCellCount must not be negative");
    }
    observedTypes = GridGrindResponseSupport.copyValues(observedTypes, "observedTypes");
  }
}
