package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;

/** High-level workbook facts returned on success. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = WorkbookSummary.Empty.class, name = "EMPTY"),
  @JsonSubTypes.Type(value = WorkbookSummary.WithSheets.class, name = "WITH_SHEETS")
})
public sealed interface WorkbookSummary permits WorkbookSummary.Empty, WorkbookSummary.WithSheets {
  /** Total sheet count after all mutations complete. */
  int sheetCount();

  /** Ordered workbook sheet names. */
  List<String> sheetNames();

  /** Count of exposed named ranges after all mutations complete. */
  int namedRangeCount();

  /** Whether the workbook is marked to recalculate formulas on open. */
  boolean forceFormulaRecalculationOnOpen();

  /** Workbook summary for a zero-sheet workbook. */
  record Empty(
      int sheetCount,
      List<String> sheetNames,
      int namedRangeCount,
      boolean forceFormulaRecalculationOnOpen)
      implements WorkbookSummary {
    public Empty {
      sheetNames =
          GridGrindResponseSupport.validateCommonWorkbookSummaryFields(
              sheetCount, sheetNames, namedRangeCount);
      if (sheetCount != 0) {
        throw new IllegalArgumentException("sheetCount must be 0 for an empty workbook");
      }
    }
  }

  /** Workbook summary for a workbook that contains one or more sheets. */
  record WithSheets(
      int sheetCount,
      List<String> sheetNames,
      String activeSheetName,
      List<String> selectedSheetNames,
      int namedRangeCount,
      boolean forceFormulaRecalculationOnOpen)
      implements WorkbookSummary {
    public WithSheets {
      sheetNames =
          GridGrindResponseSupport.validateCommonWorkbookSummaryFields(
              sheetCount, sheetNames, namedRangeCount);
      Objects.requireNonNull(activeSheetName, "activeSheetName must not be null");
      if (activeSheetName.isBlank()) {
        throw new IllegalArgumentException("activeSheetName must not be blank");
      }
      selectedSheetNames =
          GridGrindResponseSupport.copyDistinctStrings(selectedSheetNames, "selectedSheetNames");
      if (sheetCount == 0) {
        throw new IllegalArgumentException("sheetCount must be greater than 0");
      }
      if (!sheetNames.contains(activeSheetName)) {
        throw new IllegalArgumentException("activeSheetName must be present in sheetNames");
      }
      if (selectedSheetNames.isEmpty()) {
        throw new IllegalArgumentException("selectedSheetNames must not be empty");
      }
      for (String selectedSheetName : selectedSheetNames) {
        if (!sheetNames.contains(selectedSheetName)) {
          throw new IllegalArgumentException(
              "selectedSheetNames must only contain values present in sheetNames");
        }
      }
    }
  }
}
