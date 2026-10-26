package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.layout.PaneInput;
import dev.erst.gridgrind.core.model.layout.SheetPresentationInput;
import dev.erst.gridgrind.core.model.print.PrintLayoutInput;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionInput;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.Objects;
import java.util.Optional;

/** Mutation family for workbook, sheet, row, column, and layout state. */
public sealed interface WorkbookMutationAction extends MutationAction {
  /** Ensures a sheet with the given name exists, creating it if absent. */
  @Variant(
      id = "ENSURE_SHEET",
      summary = "Create the sheet if it does not already exist.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record EnsureSheet() implements WorkbookMutationAction {
    public EnsureSheet {}
  }

  /** Renames an existing sheet to a new destination name. */
  @Variant(
      id = "RENAME_SHEET",
      summary = "Rename an existing sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record RenameSheet(SheetName newSheetName) implements WorkbookMutationAction {
    public RenameSheet {
      java.util.Objects.requireNonNull(newSheetName, "newSheetName must not be null");
    }
  }

  /** Deletes an existing sheet from the workbook. */
  @Variant(
      id = "DELETE_SHEET",
      summary = "Delete an existing sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record DeleteSheet() implements WorkbookMutationAction {
    public DeleteSheet {}
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  @Variant(
      id = "MOVE_SHEET",
      summary = "Move a sheet to a zero-based workbook position.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record MoveSheet(int targetIndex) implements WorkbookMutationAction {
    public MoveSheet {
      MutationAction.Validation.requireNonNegative(targetIndex, "targetIndex");
    }
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  @Variant(
      id = "COPY_SHEET",
      summary = "Copy one sheet into a new visible, unselected sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record CopySheet(
      SheetName newSheetName,
      @ProtocolField(optional = true) @WireDefault(factory = SheetCopyPosition.AppendAtEnd.class)
          SheetCopyPosition position)
      implements WorkbookMutationAction {
    /** Copies one sheet to the end of the workbook. */
    public CopySheet(SheetName newSheetName) {
      this(newSheetName, new SheetCopyPosition.AppendAtEnd());
    }

    public CopySheet {
      java.util.Objects.requireNonNull(newSheetName, "newSheetName must not be null");
      Objects.requireNonNull(position, "position must not be null");
    }
  }

  /** Sets the active sheet and ensures it is selected. */
  @Variant(
      id = "SET_ACTIVE_SHEET",
      summary = "Set the active sheet and ensure it is selected.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetActiveSheet() implements WorkbookMutationAction {
    public SetActiveSheet {}
  }

  /** Sets the selected visible sheet set. */
  @Variant(
      id = "SET_SELECTED_SHEETS",
      summary = "Set the selected visible sheet set.",
      targetSelectors = {SheetSelector.ByNames.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetSelectedSheets() implements WorkbookMutationAction {
    public SetSelectedSheets {}
  }

  /** Sets one sheet visibility. */
  @Variant(
      id = "SET_SHEET_VISIBILITY",
      summary = "Set one sheet visibility state.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetSheetVisibility(ExcelSheetVisibility visibility) implements WorkbookMutationAction {
    public SetSheetVisibility {
      Objects.requireNonNull(visibility, "visibility must not be null");
    }
  }

  /** Enables sheet protection with the exact supported lock flags. */
  @Variant(
      id = "SET_SHEET_PROTECTION",
      summary = "Enable sheet protection with the exact supported lock flags.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetSheetProtection(
      SheetProtectionSettings protection,
      Optional<dev.erst.gridgrind.core.source.SecretReference> passwordRef)
      implements WorkbookMutationAction {
    /** Enables sheet protection without applying a password hash. */
    public SetSheetProtection(SheetProtectionSettings protection) {
      this(protection, Optional.empty());
    }

    public SetSheetProtection {
      Objects.requireNonNull(protection, "protection must not be null");
      Objects.requireNonNull(passwordRef, "passwordRef must not be null");
    }
  }

  /** Disables sheet protection entirely. */
  @Variant(
      id = "CLEAR_SHEET_PROTECTION",
      summary = "Disable sheet protection entirely.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record ClearSheetProtection() implements WorkbookMutationAction {
    public ClearSheetProtection {}
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  @Variant(
      id = "SET_WORKBOOK_PROTECTION",
      summary = "Enable workbook-level protection and optional password hashes.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetWorkbookProtection(WorkbookProtectionInput protection)
      implements WorkbookMutationAction {
    public SetWorkbookProtection {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Clears workbook-level protection and password hashes entirely. */
  @Variant(
      id = "CLEAR_WORKBOOK_PROTECTION",
      summary = "Clear workbook-level protection and stored workbook passwords.",
      targetSelectors = {WorkbookSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record ClearWorkbookProtection() implements WorkbookMutationAction {
    public ClearWorkbookProtection {}
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  @Variant(
      id = "MERGE_CELLS",
      summary = "Merge a rectangular A1-style range.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record MergeCells() implements WorkbookMutationAction {
    public MergeCells {}
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  @Variant(
      id = "UNMERGE_CELLS",
      summary = "Remove one merged region by exact range match.",
      targetSelectors = {RangeSelector.ByRange.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record UnmergeCells() implements WorkbookMutationAction {
    public UnmergeCells {}
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  @Variant(
      id = "SET_COLUMN_WIDTH",
      summary = "Set one or more column widths in Excel character units.",
      targetSelectors = {ColumnBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetColumnWidth(double widthCharacters) implements WorkbookMutationAction {
    public SetColumnWidth {
      MutationAction.Validation.requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  @Variant(
      id = "SET_ROW_HEIGHT",
      summary = "Set one or more row heights in Excel point units.",
      targetSelectors = {RowBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetRowHeight(double heightPoints) implements WorkbookMutationAction {
    public SetRowHeight {
      MutationAction.Validation.requireRowHeightPoints(heightPoints);
    }
  }

  /** Inserts one or more blank rows before the provided zero-based row index. */
  @Variant(
      id = "INSERT_ROWS",
      summary = "Insert one or more blank rows before rowIndex.",
      targetSelectors = {RowBandSelector.Insertion.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record InsertRows() implements WorkbookMutationAction {
    public InsertRows {}
  }

  /** Deletes the requested inclusive zero-based row band. */
  @Variant(
      id = "DELETE_ROWS",
      summary = "Delete one inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record DeleteRows() implements WorkbookMutationAction {
    public DeleteRows {}
  }

  /** Moves the requested inclusive zero-based row band by the provided signed delta. */
  @Variant(
      id = "SHIFT_ROWS",
      summary = "Move one inclusive zero-based row band by delta rows.",
      targetSelectors = {RowBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record ShiftRows(int delta) implements WorkbookMutationAction {
    public ShiftRows {
      MutationAction.Validation.requireNonZero(delta, "delta");
    }
  }

  /** Inserts one or more blank columns before the provided zero-based column index. */
  @Variant(
      id = "INSERT_COLUMNS",
      summary = "Insert one or more blank columns before columnIndex.",
      targetSelectors = {ColumnBandSelector.Insertion.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION,
      preconditions = {
        dev.erst.gridgrind.core.semantics.OperationPrecondition.ColumnEditsBeforeFormulaAuthoring
            .class
      })
  record InsertColumns() implements WorkbookMutationAction {
    public InsertColumns {}
  }

  /** Deletes the requested inclusive zero-based column band. */
  @Variant(
      id = "DELETE_COLUMNS",
      summary = "Delete one inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION,
      preconditions = {
        dev.erst.gridgrind.core.semantics.OperationPrecondition.ColumnEditsBeforeFormulaAuthoring
            .class
      })
  record DeleteColumns() implements WorkbookMutationAction {
    public DeleteColumns {}
  }

  /** Moves the requested inclusive zero-based column band by the provided signed delta. */
  @Variant(
      id = "SHIFT_COLUMNS",
      summary = "Move one inclusive zero-based column band by delta columns.",
      targetSelectors = {ColumnBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION,
      preconditions = {
        dev.erst.gridgrind.core.semantics.OperationPrecondition.ColumnEditsBeforeFormulaAuthoring
            .class
      })
  record ShiftColumns(int delta) implements WorkbookMutationAction {
    public ShiftColumns {
      MutationAction.Validation.requireNonZero(delta, "delta");
    }
  }

  /** Sets the hidden state for the requested inclusive zero-based row band. */
  @Variant(
      id = "SET_ROW_VISIBILITY",
      summary = "Set the hidden state for one inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetRowVisibility(boolean hidden) implements WorkbookMutationAction {}

  /** Sets the hidden state for the requested inclusive zero-based column band. */
  @Variant(
      id = "SET_COLUMN_VISIBILITY",
      summary = "Set the hidden state for one inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetColumnVisibility(boolean hidden) implements WorkbookMutationAction {}

  /** Applies one outline group to the requested inclusive zero-based row band. */
  @Variant(
      id = "GROUP_ROWS",
      summary = "Apply one outline group to an inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record GroupRows(boolean collapsed) implements WorkbookMutationAction {
    /** Creates one expanded row-group payload explicitly. */
    public static GroupRows expanded() {
      return new GroupRows(false);
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based row band. */
  @Variant(
      id = "UNGROUP_ROWS",
      summary = "Remove outline grouping from one inclusive zero-based row band.",
      targetSelectors = {RowBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record UngroupRows() implements WorkbookMutationAction {
    public UngroupRows {}
  }

  /** Applies one outline group to the requested inclusive zero-based column band. */
  @Variant(
      id = "GROUP_COLUMNS",
      summary = "Apply one outline group to an inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record GroupColumns(boolean collapsed) implements WorkbookMutationAction {
    /** Creates one expanded column-group payload explicitly. */
    public static GroupColumns expanded() {
      return new GroupColumns(false);
    }
  }

  /** Removes outline grouping from the requested inclusive zero-based column band. */
  @Variant(
      id = "UNGROUP_COLUMNS",
      summary = "Remove outline grouping from one inclusive zero-based column band.",
      targetSelectors = {ColumnBandSelector.Span.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record UngroupColumns() implements WorkbookMutationAction {
    public UngroupColumns {}
  }

  /** Applies one explicit pane state to a sheet. */
  @Variant(
      id = "SET_SHEET_PANE",
      summary = "Apply one explicit pane state to a sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetSheetPane(PaneInput pane) implements WorkbookMutationAction {
    public SetSheetPane {
      Objects.requireNonNull(pane, "pane must not be null");
    }
  }

  /** Applies one explicit zoom percentage to a sheet. */
  @Variant(
      id = "SET_SHEET_ZOOM",
      summary = "Set the sheet zoom percentage.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetSheetZoom(int zoomPercent) implements WorkbookMutationAction {
    public SetSheetZoom {
      MutationAction.Validation.requireZoomPercent(zoomPercent);
    }
  }

  /** Applies authoritative sheet-presentation state such as display flags and defaults. */
  @Variant(
      id = "SET_SHEET_PRESENTATION",
      summary = "Apply supported sheet-presentation state such as display flags and defaults.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetSheetPresentation(SheetPresentationInput presentation)
      implements WorkbookMutationAction {
    public SetSheetPresentation {
      Objects.requireNonNull(presentation, "presentation must not be null");
    }
  }

  /** Applies one authoritative supported print-layout state to a sheet. */
  @Variant(
      id = "SET_PRINT_LAYOUT",
      summary = "Apply supported print-layout state to a sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record SetPrintLayout(PrintLayoutInput printLayout) implements WorkbookMutationAction {
    public SetPrintLayout {
      Objects.requireNonNull(printLayout, "printLayout must not be null");
    }
  }

  /** Clears the supported print-layout state from a sheet. */
  @Variant(
      id = "CLEAR_PRINT_LAYOUT",
      summary = "Clear the supported print-layout state from a sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record ClearPrintLayout() implements WorkbookMutationAction {
    public ClearPrintLayout {}
  }

  /** Auto-sizes all populated columns on the sheet to fit their content. */
  @Variant(
      id = "AUTO_SIZE_COLUMNS",
      summary = "Auto-size all populated columns on the sheet to fit their content.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.WORKBOOK_MUTATION)
  record AutoSizeColumns() implements WorkbookMutationAction {
    public AutoSizeColumns {}
  }
}
