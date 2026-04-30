package dev.erst.gridgrind.contract.action;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.Objects;

/** Mutation family for workbook, sheet, row, column, and layout state. */
public sealed interface WorkbookMutationAction extends MutationAction {
  /** Ensures a sheet with the given name exists, creating it if absent. */
  record EnsureSheet() implements WorkbookMutationAction {
    public EnsureSheet {}
  }

  /** Renames an existing sheet to a new destination name. */
  record RenameSheet(String newSheetName) implements WorkbookMutationAction {
    public RenameSheet {
      MutationAction.Validation.requireSheetName(newSheetName, "newSheetName");
    }
  }

  /** Deletes an existing sheet from the workbook. */
  record DeleteSheet() implements WorkbookMutationAction {
    public DeleteSheet {}
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  record MoveSheet(int targetIndex) implements WorkbookMutationAction {
    public MoveSheet {
      MutationAction.Validation.requireNonNegative(targetIndex, "targetIndex");
    }
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  record CopySheet(String newSheetName, SheetCopyPosition position)
      implements WorkbookMutationAction {
    /** Copies one sheet to the end of the workbook. */
    public CopySheet(String newSheetName) {
      this(newSheetName, new SheetCopyPosition.AppendAtEnd());
    }

    public CopySheet {
      MutationAction.Validation.requireSheetName(newSheetName, "newSheetName");
      Objects.requireNonNull(position, "position must not be null");
    }
  }

  /** Sets the active sheet and ensures it is selected. */
  record SetActiveSheet() implements WorkbookMutationAction {
    public SetActiveSheet {}
  }

  /** Sets the selected visible sheet set. */
  record SetSelectedSheets() implements WorkbookMutationAction {
    public SetSelectedSheets {}
  }

  /** Sets one sheet visibility. */
  record SetSheetVisibility(ExcelSheetVisibility visibility) implements WorkbookMutationAction {
    public SetSheetVisibility {
      Objects.requireNonNull(visibility, "visibility must not be null");
    }
  }

  /** Enables sheet protection with the exact supported lock flags. */
  record SetSheetProtection(SheetProtectionSettings protection, String password)
      implements WorkbookMutationAction {
    /** Enables sheet protection without applying a password hash. */
    public SetSheetProtection(SheetProtectionSettings protection) {
      this(protection, null);
    }

    public SetSheetProtection {
      Objects.requireNonNull(protection, "protection must not be null");
      if (password != null && password.isBlank()) {
        throw new IllegalArgumentException("password must not be blank");
      }
    }
  }

  /** Disables sheet protection entirely. */
  record ClearSheetProtection() implements WorkbookMutationAction {
    public ClearSheetProtection {}
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  record SetWorkbookProtection(WorkbookProtectionInput protection)
      implements WorkbookMutationAction {
    public SetWorkbookProtection {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Clears workbook-level protection and password hashes entirely. */
  record ClearWorkbookProtection() implements WorkbookMutationAction {
    public ClearWorkbookProtection {}
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  record MergeCells() implements WorkbookMutationAction {
    public MergeCells {}
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  record UnmergeCells() implements WorkbookMutationAction {
    public UnmergeCells {}
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  record SetColumnWidth(double widthCharacters) implements WorkbookMutationAction {
    public SetColumnWidth {
      MutationAction.Validation.requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  record SetRowHeight(double heightPoints) implements WorkbookMutationAction {
    public SetRowHeight {
      MutationAction.Validation.requireRowHeightPoints(heightPoints);
    }
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  record InsertRows() implements WorkbookMutationAction {
    public InsertRows {}
  }

  /** Deletes the requested inclusive zero-based row band. */
  record DeleteRows() implements WorkbookMutationAction {
    public DeleteRows {}
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  record ShiftRows(int delta) implements WorkbookMutationAction {
    public ShiftRows {
      MutationAction.Validation.requireNonZero(delta, "delta");
    }
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  record InsertColumns() implements WorkbookMutationAction {
    public InsertColumns {}
  }

  /** Deletes the requested inclusive zero-based column band. */
  record DeleteColumns() implements WorkbookMutationAction {
    public DeleteColumns {}
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  record ShiftColumns(int delta) implements WorkbookMutationAction {
    public ShiftColumns {
      MutationAction.Validation.requireNonZero(delta, "delta");
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  record SetRowVisibility(boolean hidden) implements WorkbookMutationAction {}

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  record SetColumnVisibility(boolean hidden) implements WorkbookMutationAction {}

  /** Applies one outline group to the requested inclusive zero-based row band. */
  record GroupRows(boolean collapsed) implements WorkbookMutationAction {
    /** Creates one expanded row-group payload explicitly. */
    public static GroupRows expanded() {
      return new GroupRows(false);
    }

    /** Reads one row-group payload with explicit collapse state. */
    @JsonCreator
    public GroupRows(@JsonProperty("collapsed") Boolean collapsed) {
      this(Objects.requireNonNull(collapsed, "collapsed must not be null").booleanValue());
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  record UngroupRows() implements WorkbookMutationAction {
    public UngroupRows {}
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  record GroupColumns(boolean collapsed) implements WorkbookMutationAction {
    /** Creates one expanded column-group payload explicitly. */
    public static GroupColumns expanded() {
      return new GroupColumns(false);
    }

    /** Reads one column-group payload with explicit collapse state. */
    @JsonCreator
    public GroupColumns(@JsonProperty("collapsed") Boolean collapsed) {
      this(Objects.requireNonNull(collapsed, "collapsed must not be null").booleanValue());
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  record UngroupColumns() implements WorkbookMutationAction {
    public UngroupColumns {}
  }

  /** Applies one explicit pane state to a sheet. */
  record SetSheetPane(PaneInput pane) implements WorkbookMutationAction {
    public SetSheetPane {
      Objects.requireNonNull(pane, "pane must not be null");
    }
  }

  /** Applies one explicit zoom percentage to a sheet. */
  record SetSheetZoom(int zoomPercent) implements WorkbookMutationAction {
    public SetSheetZoom {
      MutationAction.Validation.requireZoomPercent(zoomPercent);
    }
  }

  /** Applies authoritative sheet-presentation state such as display flags and defaults. */
  record SetSheetPresentation(SheetPresentationInput presentation)
      implements WorkbookMutationAction {
    public SetSheetPresentation {
      Objects.requireNonNull(presentation, "presentation must not be null");
    }
  }

  /** Applies one authoritative supported print-layout state to a sheet. */
  record SetPrintLayout(PrintLayoutInput printLayout) implements WorkbookMutationAction {
    public SetPrintLayout {
      Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Clears the supported print-layout state from a sheet. */
  record ClearPrintLayout() implements WorkbookMutationAction {
    public ClearPrintLayout {}
  }

  /** Auto-sizes all populated columns on the sheet to fit their content. */
  record AutoSizeColumns() implements WorkbookMutationAction {
    public AutoSizeColumns {}
  }
}
