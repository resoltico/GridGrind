package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.List;
import java.util.Objects;

/** High-level workbook facts returned on success. */
@Union(discriminator = "type")
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
  @Variant(id = "EMPTY", summary = "Workbook summary for a zero-sheet workbook.")
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
  @Variant(
      id = "WITH_SHEETS",
      summary = "Workbook summary for a workbook that contains one or more sheets.")
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
