package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import dev.erst.gridgrind.excel.foundation.ExcelIndexDisplay;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
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
  record SheetSummaryResult(String stepId, SheetSummary sheet) implements WorkbookSheetResult {
    public SheetSummaryResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Returns factual array-formula groups across the selected sheets. */
  record ArrayFormulasResult(String stepId, List<ExcelArrayFormulaSnapshot> arrayFormulas)
      implements WorkbookSheetResult {
    public ArrayFormulasResult {
      stepId = requireNonBlank(stepId, "stepId");
      arrayFormulas = copyValues(arrayFormulas, "arrayFormulas");
    }
  }

  /** Returns exact cell snapshots for one sheet. */
  record CellsResult(String stepId, String sheetName, List<ExcelCellSnapshot> cells)
      implements WorkbookSheetResult {
    public CellsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      cells = copyValues(cells, "cells");
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  record WindowResult(String stepId, Window window) implements WorkbookSheetResult {
    public WindowResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(window, "window must not be null");
    }
  }

  /** Returns every merged region defined on one sheet. */
  record MergedRegionsResult(String stepId, String sheetName, List<MergedRegion> mergedRegions)
      implements WorkbookSheetResult {
    public MergedRegionsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      mergedRegions = copyValues(mergedRegions, "mergedRegions");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  record HyperlinksResult(String stepId, String sheetName, List<CellHyperlink> hyperlinks)
      implements WorkbookSheetResult {
    public HyperlinksResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      hyperlinks = copyValues(hyperlinks, "hyperlinks");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  record CommentsResult(String stepId, String sheetName, List<CellComment> comments)
      implements WorkbookSheetResult {
    public CommentsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      comments = copyValues(comments, "comments");
    }
  }

  /** Returns layout metadata such as panes, zoom, and visible sizing. */
  record SheetLayoutResult(String stepId, SheetLayout layout) implements WorkbookSheetResult {
    public SheetLayoutResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns supported print-layout metadata for one sheet. */
  record PrintLayoutResult(String stepId, String sheetName, ExcelPrintLayoutSnapshot printLayout)
      implements WorkbookSheetResult {
    public PrintLayoutResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Structural summary facts for one sheet. */
  record SheetSummary(
      String sheetName,
      ExcelSheetVisibility visibility,
      SheetProtection protection,
      int physicalRowCount,
      int lastRowIndex,
      int lastColumnIndex) {
    public SheetSummary {
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(visibility, "visibility must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
      if (physicalRowCount < 0) {
        throw new IllegalArgumentException("physicalRowCount must not be negative");
      }
      if (lastRowIndex < -1) {
        throw new IllegalArgumentException(
            "lastRowIndex "
                + lastRowIndex
                + " must be greater than or equal to -1; empty sheets report -1");
      }
      if (lastColumnIndex < -1) {
        throw new IllegalArgumentException(
            "lastColumnIndex "
                + lastColumnIndex
                + " must be greater than or equal to -1; empty sheets report -1");
      }
    }
  }

  /** Captures whether a sheet is protected and, if so, with which supported lock flags. */
  sealed interface SheetProtection permits SheetProtection.Unprotected, SheetProtection.Protected {

    /** Sheet protection is disabled. */
    record Unprotected() implements SheetProtection {}

    /** Sheet protection is enabled with the reported supported lock flags. */
    record Protected(ExcelSheetProtectionSettings settings) implements SheetProtection {
      public Protected {
        Objects.requireNonNull(settings, "settings must not be null");
      }
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
  record WindowRow(int rowIndex, List<ExcelCellSnapshot> cells) {
    public WindowRow {
      if (rowIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("rowIndex", rowIndex));
      }
      cells = copyValues(cells, "cells");
    }
  }

  /** One merged region captured from a sheet. */
  record MergedRegion(String range) {
    public MergedRegion {
      range = requireNonBlank(range, "range");
    }
  }

  /** Hyperlink metadata associated with one concrete cell address. */
  record CellHyperlink(String address, ExcelHyperlink hyperlink) {
    public CellHyperlink {
      address = requireNonBlank(address, "address");
      Objects.requireNonNull(hyperlink, "hyperlink must not be null");
    }
  }

  /** Comment metadata associated with one concrete cell address. */
  record CellComment(String address, ExcelCommentSnapshot comment) {
    public CellComment {
      address = requireNonBlank(address, "address");
      Objects.requireNonNull(comment, "comment must not be null");
    }
  }

  /** Layout metadata such as pane state, zoom, and visible row and column sizing for one sheet. */
  record SheetLayout(
      String sheetName,
      ExcelSheetPane pane,
      int zoomPercent,
      ExcelSheetPresentationSnapshot presentation,
      List<ColumnLayout> columns,
      List<RowLayout> rows) {
    public SheetLayout {
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(pane, "pane must not be null");
      ExcelSheetViewSupport.requireZoomPercent(zoomPercent);
      Objects.requireNonNull(presentation, "presentation must not be null");
      columns = copyValues(columns, "columns");
      rows = copyValues(rows, "rows");
    }
  }

  /** Width metadata for one sheet column. */
  record ColumnLayout(
      int columnIndex,
      double widthCharacters,
      boolean hidden,
      int outlineLevel,
      boolean collapsed) {
    public ColumnLayout {
      if (columnIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("columnIndex", columnIndex));
      }
      if (!Double.isFinite(widthCharacters) || widthCharacters <= 0.0d) {
        throw new IllegalArgumentException("widthCharacters must be finite and greater than 0");
      }
      if (outlineLevel < 0) {
        throw new IllegalArgumentException("outlineLevel must not be negative");
      }
    }
  }

  /** Height metadata for one sheet row. */
  record RowLayout(
      int rowIndex, double heightPoints, boolean hidden, int outlineLevel, boolean collapsed) {
    public RowLayout {
      if (rowIndex < 0) {
        throw new IllegalArgumentException(
            ExcelIndexDisplay.mustNotBeNegative("rowIndex", rowIndex));
      }
      if (!Double.isFinite(heightPoints) || heightPoints <= 0.0d) {
        throw new IllegalArgumentException("heightPoints must be finite and greater than 0");
      }
      if (outlineLevel < 0) {
        throw new IllegalArgumentException("outlineLevel must not be negative");
      }
    }
  }
}
