package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.vocabulary.ExcelIndexDisplay;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Objects;

/** Sheet-scoped fact results and sheet layout payloads. */
public sealed interface WorkbookSheetResult extends WorkbookReadIntrospectionResult
    permits WorkbookSheetResult.SheetSummaryResult,
        WorkbookSheetResult.ArrayFormulasResult,
        WorkbookSheetResult.CellsResult,
        WorkbookSheetResult.WindowResult,
        WorkbookSheetResult.MergedRegionsResult,
        WorkbookSheetResult.HyperlinksResult,
        WorkbookSheetResult.CommentsResult,
        WorkbookSheetResult.SheetLayoutResult,
        WorkbookSheetResult.PrintLayoutResult {

  /** Returns summary facts for one sheet. */
  record SheetSummaryResult(StepId stepId, SheetSummaryReport sheet)
      implements WorkbookSheetResult {
    public SheetSummaryResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Returns factual array-formula groups across the selected sheets. */
  record ArrayFormulasResult(StepId stepId, List<ArrayFormulaReport> arrayFormulas)
      implements WorkbookSheetResult {
    public ArrayFormulasResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      arrayFormulas = copyValues(arrayFormulas, "arrayFormulas");
    }
  }

  /** Returns exact cell snapshots for one sheet. */
  record CellsResult(
      StepId stepId,
      String sheetName,
      List<CellReport> cells,
      CellReadProjection projection,
      boolean date1904)
      implements WorkbookSheetResult {
    public CellsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      cells = copyValues(cells, "cells");
      Objects.requireNonNull(projection, "projection must not be null");
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  record WindowResult(
      StepId stepId,
      Window window,
      CellReadProjection projection,
      boolean includeBlanks,
      boolean date1904)
      implements WorkbookSheetResult {
    public WindowResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(window, "window must not be null");
      Objects.requireNonNull(projection, "projection must not be null");
    }
  }

  /** Returns every merged region defined on one sheet. */
  record MergedRegionsResult(
      StepId stepId, String sheetName, List<MergedRegionReport> mergedRegions)
      implements WorkbookSheetResult {
    public MergedRegionsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      mergedRegions = copyValues(mergedRegions, "mergedRegions");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record HyperlinksResult(StepId stepId, String sheetName, List<CellHyperlinkReport> hyperlinks)
      implements WorkbookSheetResult {
    public HyperlinksResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      hyperlinks = copyValues(hyperlinks, "hyperlinks");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record CommentsResult(StepId stepId, String sheetName, List<CellCommentReport> comments)
      implements WorkbookSheetResult {
    public CommentsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      comments = copyValues(comments, "comments");
    }
  }

  /** Returns layout metadata such as panes, zoom, and visible sizing. */
  record SheetLayoutResult(StepId stepId, SheetLayoutReport layout) implements WorkbookSheetResult {
    public SheetLayoutResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns supported print-layout metadata for one sheet. */
  record PrintLayoutResult(StepId stepId, String sheetName, PrintLayoutReport printLayout)
      implements WorkbookSheetResult {
    public PrintLayoutResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Rectangular window of cell snapshots anchored at one top-left address. */
  record Window(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      List<WindowRow> rows) {
    public Window {
      sheetName = requireNonBlank(sheetName, "sheetName");
      topLeftAddress = requireNonBlank(topLeftAddress, "topLeftAddress");
      if (rowCount <= 0) {
        throw new IllegalArgumentException("rowCount must be greater than 0");
      }
      if (columnCount <= 0) {
        throw new IllegalArgumentException("columnCount must be greater than 0");
      }
      rows = copyValues(rows, "rows");
    }
  }

  /** One row inside a rectangular window of cell snapshots. */
  record WindowRow(int rowIndex, List<CellReport> cells) {
    public WindowRow {
      if (rowIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("rowIndex", rowIndex));
      }
      cells = copyValues(cells, "cells");
    }
  }
}
