package dev.erst.gridgrind.contract.action;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.Objects;
import java.util.Optional;

/** Mutation family for workbook, sheet, row, column, and layout state. */
public sealed interface WorkbookMutationAction extends MutationAction {
  /** Ensures a sheet with the given name exists, creating it if absent. */
  @ProtocolTypeMetadata(
      id = "ENSURE_SHEET",
      summary = "Create the sheet if it does not already exist.",
      targetSelectors = {SheetSelector.ByName.class})
  record EnsureSheet() implements WorkbookMutationAction {
    public EnsureSheet {}
  }

  /** Renames an existing sheet to a new destination name. */
  @ProtocolTypeMetadata(
      id = "RENAME_SHEET",
      summary = "Rename an existing sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record RenameSheet(String newSheetName) implements WorkbookMutationAction {
    public RenameSheet {
      MutationAction.Validation.requireSheetName(newSheetName, "newSheetName");
    }
  }

  /** Deletes an existing sheet from the workbook. */
  @ProtocolTypeMetadata(
      id = "DELETE_SHEET",
      summary = "Delete an existing sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record DeleteSheet() implements WorkbookMutationAction {
    public DeleteSheet {}
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  @ProtocolTypeMetadata(
      id = "MOVE_SHEET",
      summary = "Move a sheet to a zero-based workbook position.",
      targetSelectors = {SheetSelector.ByName.class})
  record MoveSheet(int targetIndex) implements WorkbookMutationAction {
    public MoveSheet {
      MutationAction.Validation.requireNonNegative(targetIndex, "targetIndex");
    }
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  @ProtocolTypeMetadata(
      id = "COPY_SHEET",
      summary = "Copy one sheet into a new visible, unselected sheet.",
      optionalFields = {"position"},
      targetSelectors = {SheetSelector.ByName.class})
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
  @ProtocolTypeMetadata(
      id = "SET_ACTIVE_SHEET",
      summary = "Set the active sheet and ensure it is selected.",
      targetSelectors = {SheetSelector.ByName.class})
  record SetActiveSheet() implements WorkbookMutationAction {
    public SetActiveSheet {}
  }

  /** Sets the selected visible sheet set. */
  @ProtocolTypeMetadata(
      id = "SET_SELECTED_SHEETS",
      summary = "Set the selected visible sheet set.",
      targetSelectors = {SheetSelector.ByNames.class})
  record SetSelectedSheets() implements WorkbookMutationAction {
    public SetSelectedSheets {}
  }

  /** Sets one sheet visibility. */
  @ProtocolTypeMetadata(
      id = "SET_SHEET_VISIBILITY",
      summary = "Set one sheet visibility state.",
      targetSelectors = {SheetSelector.ByName.class})
  record SetSheetVisibility(ExcelSheetVisibility visibility) implements WorkbookMutationAction {
    public SetSheetVisibility {
      Objects.requireNonNull(visibility, "visibility must not be null");
    }
  }

  /** Enables sheet protection with the exact supported lock flags. */
  @ProtocolTypeMetadata(
      id = "SET_SHEET_PROTECTION",
      summary = "Enable sheet protection with the exact supported lock flags.",
      optionalFields = {"password"},
      targetSelectors = {SheetSelector.ByName.class})
  record SetSheetProtection(
      SheetProtectionSettings protection,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> password)
      implements WorkbookMutationAction {
    /** Enables sheet protection without applying a password hash. */
    public SetSheetProtection(SheetProtectionSettings protection) {
      this(protection, Optional.empty());
    }

    /** Enables sheet protection with one explicit password string. */
    public SetSheetProtection(SheetProtectionSettings protection, String password) {
      this(protection, Optional.ofNullable(password));
    }

    public SetSheetProtection {
      Objects.requireNonNull(protection, "protection must not be null");
      Objects.requireNonNull(password, "password must not be null");
      if (password.isPresent() && password.orElseThrow().isBlank()) {
        throw new IllegalArgumentException("password must not be blank");
      }
    }
  }

  /** Disables sheet protection entirely. */
  @ProtocolTypeMetadata(
      id = "CLEAR_SHEET_PROTECTION",
      summary = "Disable sheet protection entirely.",
      targetSelectors = {SheetSelector.ByName.class})
  record ClearSheetProtection() implements WorkbookMutationAction {
    public ClearSheetProtection {}
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  @ProtocolTypeMetadata(
      id = "SET_WORKBOOK_PROTECTION",
      summary = "Enable workbook-level protection and optional password hashes.",
      targetSelectors = {WorkbookSelector.class})
  record SetWorkbookProtection(WorkbookProtectionInput protection)
      implements WorkbookMutationAction {
    public SetWorkbookProtection {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Clears workbook-level protection and password hashes entirely. */
  @ProtocolTypeMetadata(
      id = "CLEAR_WORKBOOK_PROTECTION",
      summary = "Clear workbook-level protection and stored workbook passwords.",
      targetSelectors = {WorkbookSelector.class})
  record ClearWorkbookProtection() implements WorkbookMutationAction {
    public ClearWorkbookProtection {}
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  @ProtocolTypeMetadata(
      id = "MERGE_CELLS",
      summary = "Merge a rectangular A1-style range.",
      targetSelectors = {RangeSelector.ByRange.class})
  record MergeCells() implements WorkbookMutationAction {
    public MergeCells {}
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  @ProtocolTypeMetadata(
      id = "UNMERGE_CELLS",
      summary = "Remove one merged region by exact range match.",
      targetSelectors = {RangeSelector.ByRange.class})
  record UnmergeCells() implements WorkbookMutationAction {
    public UnmergeCells {}
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  @ProtocolTypeMetadata(
      id = "SET_COLUMN_WIDTH",
      summary = "Set one or more column widths in Excel character units.",
      targetSelectors = {ColumnBandSelector.Span.class})
  record SetColumnWidth(double widthCharacters) implements WorkbookMutationAction {
    public SetColumnWidth {
      MutationAction.Validation.requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  @ProtocolTypeMetadata(
      id = "SET_ROW_HEIGHT",
      summary = "Set one or more row heights in Excel point units.",
      targetSelectors = {RowBandSelector.Span.class})
  record SetRowHeight(double heightPoints) implements WorkbookMutationAction {
    public SetRowHeight {
      MutationAction.Validation.requireRowHeightPoints(heightPoints);
    }
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  @ProtocolTypeMetadata(
      id = "INSERT_ROWS",
      summary = "Insert one or more blank rows before rowIndex.",
      targetSelectors = {RowBandSelector.Insertion.class})
  record InsertRows() implements WorkbookMutationAction {
    public InsertRows {}
  }

  /** Deletes the requested inclusive zero-based row band. */
  @ProtocolTypeMetadata(
      id = "DELETE_ROWS",
      summary = "Delete one inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class})
  record DeleteRows() implements WorkbookMutationAction {
    public DeleteRows {}
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  @ProtocolTypeMetadata(
      id = "SHIFT_ROWS",
      summary = "Move one inclusive zero-based row band by delta rows.",
      targetSelectors = {RowBandSelector.Span.class})
  record ShiftRows(int delta) implements WorkbookMutationAction {
    public ShiftRows {
      MutationAction.Validation.requireNonZero(delta, "delta");
    }
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  @ProtocolTypeMetadata(
      id = "INSERT_COLUMNS",
      summary = "Insert one or more blank columns before columnIndex.",
      targetSelectors = {ColumnBandSelector.Insertion.class})
  record InsertColumns() implements WorkbookMutationAction {
    public InsertColumns {}
  }

  /** Deletes the requested inclusive zero-based column band. */
  @ProtocolTypeMetadata(
      id = "DELETE_COLUMNS",
      summary = "Delete one inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class})
  record DeleteColumns() implements WorkbookMutationAction {
    public DeleteColumns {}
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  @ProtocolTypeMetadata(
      id = "SHIFT_COLUMNS",
      summary = "Move one inclusive zero-based column band by delta columns.",
      targetSelectors = {ColumnBandSelector.Span.class})
  record ShiftColumns(int delta) implements WorkbookMutationAction {
    public ShiftColumns {
      MutationAction.Validation.requireNonZero(delta, "delta");
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  @ProtocolTypeMetadata(
      id = "SET_ROW_VISIBILITY",
      summary = "Set the hidden state for one inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class})
  record SetRowVisibility(boolean hidden) implements WorkbookMutationAction {}

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  @ProtocolTypeMetadata(
      id = "SET_COLUMN_VISIBILITY",
      summary = "Set the hidden state for one inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class})
  record SetColumnVisibility(boolean hidden) implements WorkbookMutationAction {}

  /** Applies one outline group to the requested inclusive zero-based row band. */
  @ProtocolTypeMetadata(
      id = "GROUP_ROWS",
      summary = "Apply one outline group to an inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class})
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
  @ProtocolTypeMetadata(
      id = "UNGROUP_ROWS",
      summary = "Remove outline grouping from one inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class})
  record UngroupRows() implements WorkbookMutationAction {
    public UngroupRows {}
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  @ProtocolTypeMetadata(
      id = "GROUP_COLUMNS",
      summary = "Apply one outline group to an inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class})
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
  @ProtocolTypeMetadata(
      id = "UNGROUP_COLUMNS",
      summary = "Remove outline grouping from one inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class})
  record UngroupColumns() implements WorkbookMutationAction {
    public UngroupColumns {}
  }

  /** Applies one explicit pane state to a sheet. */
  @ProtocolTypeMetadata(
      id = "SET_SHEET_PANE",
      summary = "Apply one explicit pane state to a sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record SetSheetPane(PaneInput pane) implements WorkbookMutationAction {
    public SetSheetPane {
      Objects.requireNonNull(pane, "pane must not be null");
    }
  }

  /** Applies one explicit zoom percentage to a sheet. */
  @ProtocolTypeMetadata(
      id = "SET_SHEET_ZOOM",
      summary = "Set the sheet zoom percentage.",
      targetSelectors = {SheetSelector.ByName.class})
  record SetSheetZoom(int zoomPercent) implements WorkbookMutationAction {
    public SetSheetZoom {
      MutationAction.Validation.requireZoomPercent(zoomPercent);
    }
  }

  /** Applies authoritative sheet-presentation state such as display flags and defaults. */
  @ProtocolTypeMetadata(
      id = "SET_SHEET_PRESENTATION",
      summary = "Apply supported sheet-presentation state such as display flags and defaults.",
      targetSelectors = {SheetSelector.ByName.class})
  record SetSheetPresentation(SheetPresentationInput presentation)
      implements WorkbookMutationAction {
    public SetSheetPresentation {
      Objects.requireNonNull(presentation, "presentation must not be null");
    }
  }

  /** Applies one authoritative supported print-layout state to a sheet. */
  @ProtocolTypeMetadata(
      id = "SET_PRINT_LAYOUT",
      summary = "Apply supported print-layout state to a sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record SetPrintLayout(PrintLayoutInput printLayout) implements WorkbookMutationAction {
    public SetPrintLayout {
      Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Clears the supported print-layout state from a sheet. */
  @ProtocolTypeMetadata(
      id = "CLEAR_PRINT_LAYOUT",
      summary = "Clear the supported print-layout state from a sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record ClearPrintLayout() implements WorkbookMutationAction {
    public ClearPrintLayout {}
  }

  /** Auto-sizes all populated columns on the sheet to fit their content. */
  @ProtocolTypeMetadata(
      id = "AUTO_SIZE_COLUMNS",
      summary = "Auto-size all populated columns on the sheet to fit their content.",
      targetSelectors = {SheetSelector.ByName.class})
  record AutoSizeColumns() implements WorkbookMutationAction {
    public AutoSizeColumns {}
  }
}
