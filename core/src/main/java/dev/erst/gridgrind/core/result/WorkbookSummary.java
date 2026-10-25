package dev.erst.gridgrind.core.result;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
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
  List<SheetLabel> sheetNames();

  /** Count of exposed named ranges after all mutations complete. */
  int namedRangeCount();

  /** Whether the workbook is marked to recalculate formulas on open. */
  boolean forceFormulaRecalculationOnOpen();

  /** Workbook summary for a zero-sheet workbook. */
  record Empty(
      int sheetCount,
      List<SheetLabel> sheetNames,
      int namedRangeCount,
      boolean forceFormulaRecalculationOnOpen)
      implements WorkbookSummary {
    public Empty {
      sheetNames =
          WorkbookResultSupport.validateCommonWorkbookSummaryFields(
              sheetCount, sheetNames, namedRangeCount);
      if (sheetCount != 0) {
        throw new IllegalArgumentException("sheetCount must be 0 for an empty workbook");
      }
    }
  }

  /** Workbook summary for a workbook that contains one or more sheets. */
  record WithSheets(
      int sheetCount,
      List<SheetLabel> sheetNames,
      SheetLabel activeSheetName,
      List<SheetLabel> selectedSheetNames,
      int namedRangeCount,
      boolean forceFormulaRecalculationOnOpen)
      implements WorkbookSummary {
    public WithSheets {
      sheetNames =
          WorkbookResultSupport.validateCommonWorkbookSummaryFields(
              sheetCount, sheetNames, namedRangeCount);
      Objects.requireNonNull(activeSheetName, "activeSheetName must not be null");
      selectedSheetNames =
          WorkbookResultSupport.copyDistinctValues(selectedSheetNames, "selectedSheetNames");
      if (sheetCount == 0) {
        throw new IllegalArgumentException("sheetCount must be greater than 0");
      }
      if (!sheetNames.contains(activeSheetName)) {
        throw new IllegalArgumentException("activeSheetName must be present in sheetNames");
      }
      if (selectedSheetNames.isEmpty()) {
        throw new IllegalArgumentException("selectedSheetNames must not be empty");
      }
      for (SheetLabel selectedSheetName : selectedSheetNames) {
        if (!sheetNames.contains(selectedSheetName)) {
          throw new IllegalArgumentException(
              "selectedSheetNames must only contain values present in sheetNames");
        }
      }
    }
  }
}
