package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.result.TypeCountReport;
import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** One inferred schema column with header text and observed value-type counts. */
public record SchemaColumnReport(
    int columnIndex,
    String columnAddress,
    String headerDisplayValue,
    int populatedCellCount,
    int blankCellCount,
    List<TypeCountReport> observedTypes,
    Optional<String> dominantType) {
  public SchemaColumnReport {
    Objects.requireNonNull(dominantType, "dominantType must not be null");
    if (columnIndex < 0) {
      throw new IllegalArgumentException("columnIndex must not be negative");
    }
    Objects.requireNonNull(columnAddress, "columnAddress must not be null");
    Objects.requireNonNull(headerDisplayValue, "headerDisplayValue must not be null");
    TextValidation.requireNonBlank(columnAddress, "columnAddress");
    if (populatedCellCount < 0) {
      throw new IllegalArgumentException("populatedCellCount must not be negative");
    }
    if (blankCellCount < 0) {
      throw new IllegalArgumentException("blankCellCount must not be negative");
    }
    observedTypes = WorkbookResultSupport.copyValues(observedTypes, "observedTypes");
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
    dominantType.ifPresent(
        type -> {
          TypeCountReport.requireSupportedType(type, "dominantType");
          if (!observedTypeNames.contains(type)) {
            throw new IllegalArgumentException(
                "dominantType must be omitted or match one observedTypes entry");
          }
        });
  }
}
