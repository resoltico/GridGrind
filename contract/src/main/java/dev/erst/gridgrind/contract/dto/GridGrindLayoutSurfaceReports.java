package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Window, layout, hyperlink, and comment surface reports. */
public interface GridGrindLayoutSurfaceReports {
  /** Rectangular window of cells anchored at one top-left address. */
  record WindowReport(
      String sheetName,
      String topLeftAddress,
      int rowCount,
      int columnCount,
      List<WindowRowReport> rows) {
    public WindowReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (topLeftAddress.isBlank()) {
        throw new IllegalArgumentException("topLeftAddress must not be blank");
      }
      if (rowCount <= 0) {
        throw new IllegalArgumentException("rowCount must be greater than 0");
      }
      if (columnCount <= 0) {
        throw new IllegalArgumentException("columnCount must be greater than 0");
      }
      rows = GridGrindResponseSupport.copyValues(rows, "rows");
    }
  }

  /** One row inside a rectangular window of cell snapshots. */
  record WindowRowReport(int rowIndex, List<CellReport> cells) {
    public WindowRowReport {
      cells = GridGrindResponseSupport.copyValues(cells, "cells");
    }
  }

  /** One merged region captured from a sheet. */
  record MergedRegionReport(String range) {
    public MergedRegionReport {
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Hyperlink metadata associated with one concrete cell address. */
  record CellHyperlinkReport(String address, HyperlinkTarget hyperlink) {
    public CellHyperlinkReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(hyperlink, "hyperlink must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Comment metadata associated with one concrete cell address. */
  record CellCommentReport(String address, GridGrindWorkbookSurfaceReports.CommentReport comment) {
    public CellCommentReport {
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(comment, "comment must not be null");
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Layout metadata such as pane state, zoom, and visible row and column sizing for one sheet. */
  record SheetLayoutReport(
      String sheetName,
      PaneReport pane,
      int zoomPercent,
      SheetPresentationReport presentation,
      List<ColumnLayoutReport> columns,
      List<RowLayoutReport> rows) {
    public SheetLayoutReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(pane, "pane must not be null");
      Objects.requireNonNull(presentation, "presentation must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.requireZoomPercent(
          zoomPercent, "zoomPercent");
      columns = GridGrindResponseSupport.copyValues(columns, "columns");
      rows = GridGrindResponseSupport.copyValues(rows, "rows");
    }
  }

  /** Width metadata for one sheet column. */
  record ColumnLayoutReport(
      int columnIndex,
      double widthCharacters,
      boolean hidden,
      int outlineLevel,
      boolean collapsed) {
    public ColumnLayoutReport {
      if (columnIndex < 0) {
        throw new IllegalArgumentException("columnIndex must not be negative");
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
  record RowLayoutReport(
      int rowIndex, double heightPoints, boolean hidden, int outlineLevel, boolean collapsed) {
    public RowLayoutReport {
      if (rowIndex < 0) {
        throw new IllegalArgumentException("rowIndex must not be negative");
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
