package dev.erst.gridgrind.contract.dto;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
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
    Set<String> observedTypeNames = new LinkedHashSet<>();
    int observedTypeCountTotal = 0;
    for (TypeCountReport observedType : observedTypes) {
      if (!observedTypeNames.add(observedType.type())) {
        throw new IllegalArgumentException(
            "observedTypes must not contain duplicate type " + observedType.type());
      }
      observedTypeCountTotal += observedType.count();
    }
    if (observedTypeCountTotal != populatedCellCount) {
      throw new IllegalArgumentException("observedTypes counts must sum to populatedCellCount");
    }
    if (dominantType != null) {
      TypeCountReport.requireSupportedType(dominantType, "dominantType");
      if (!observedTypeNames.contains(dominantType)) {
        throw new IllegalArgumentException(
            "dominantType must be omitted or match one observedTypes entry");
      }
    }
  }
}
