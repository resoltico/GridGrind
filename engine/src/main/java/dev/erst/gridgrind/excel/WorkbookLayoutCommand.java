package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Sheet layout commands that control panes, zoom, presentation, and print layout. */
public sealed interface WorkbookLayoutCommand extends WorkbookCommand
    permits WorkbookLayoutCommand.SetSheetPane,
        WorkbookLayoutCommand.SetSheetZoom,
        WorkbookLayoutCommand.SetSheetPresentation,
        WorkbookLayoutCommand.SetPrintLayout,
        WorkbookLayoutCommand.ClearPrintLayout,
        WorkbookLayoutCommand.AutoSizeColumns {

  /** Applies one explicit pane state to a sheet. */
  record SetSheetPane(String sheetName, ExcelSheetPane pane) implements WorkbookLayoutCommand {
    public SetSheetPane {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(pane, "pane must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Applies one explicit zoom percentage to a sheet. */
  record SetSheetZoom(String sheetName, int zoomPercent) implements WorkbookLayoutCommand {
    public SetSheetZoom {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      ExcelSheetViewSupport.requireZoomPercent(zoomPercent);
    }
  }

  /** Applies authoritative sheet-presentation state such as display flags and defaults. */
  record SetSheetPresentation(String sheetName, ExcelSheetPresentation presentation)
      implements WorkbookLayoutCommand {
    public SetSheetPresentation {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(presentation, "presentation must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Applies one authoritative supported print-layout state to a sheet. */
  record SetPrintLayout(String sheetName, ExcelPrintLayout printLayout)
      implements WorkbookLayoutCommand {
    public SetPrintLayout {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(printLayout, "printLayout must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Clears the supported print-layout state from a sheet. */
  record ClearPrintLayout(String sheetName) implements WorkbookLayoutCommand {
    public ClearPrintLayout {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Auto-sizes all populated columns on the named sheet. */
  record AutoSizeColumns(String sheetName) implements WorkbookLayoutCommand {
    public AutoSizeColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }
}
