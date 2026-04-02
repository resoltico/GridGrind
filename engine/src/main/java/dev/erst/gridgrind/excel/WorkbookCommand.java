package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/**
 * Workbook-core commands that can be executed deterministically against an {@link ExcelWorkbook}.
 */
public sealed interface WorkbookCommand
    permits WorkbookCommand.CreateSheet,
        WorkbookCommand.RenameSheet,
        WorkbookCommand.DeleteSheet,
        WorkbookCommand.MoveSheet,
        WorkbookCommand.CopySheet,
        WorkbookCommand.SetActiveSheet,
        WorkbookCommand.SetSelectedSheets,
        WorkbookCommand.SetSheetVisibility,
        WorkbookCommand.SetSheetProtection,
        WorkbookCommand.ClearSheetProtection,
        WorkbookCommand.MergeCells,
        WorkbookCommand.UnmergeCells,
        WorkbookCommand.SetColumnWidth,
        WorkbookCommand.SetRowHeight,
        WorkbookCommand.SetSheetPane,
        WorkbookCommand.SetSheetZoom,
        WorkbookCommand.SetPrintLayout,
        WorkbookCommand.ClearPrintLayout,
        WorkbookCommand.SetCell,
        WorkbookCommand.SetRange,
        WorkbookCommand.ClearRange,
        WorkbookCommand.SetHyperlink,
        WorkbookCommand.ClearHyperlink,
        WorkbookCommand.SetComment,
        WorkbookCommand.ClearComment,
        WorkbookCommand.ApplyStyle,
        WorkbookCommand.SetDataValidation,
        WorkbookCommand.ClearDataValidations,
        WorkbookCommand.SetConditionalFormatting,
        WorkbookCommand.ClearConditionalFormatting,
        WorkbookCommand.SetAutofilter,
        WorkbookCommand.ClearAutofilter,
        WorkbookCommand.SetTable,
        WorkbookCommand.DeleteTable,
        WorkbookCommand.SetNamedRange,
        WorkbookCommand.DeleteNamedRange,
        WorkbookCommand.AppendRow,
        WorkbookCommand.AutoSizeColumns,
        WorkbookCommand.EvaluateAllFormulas,
        WorkbookCommand.ForceFormulaRecalculationOnOpen {

  record CreateSheet(String sheetName) implements WorkbookCommand {
    public CreateSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Renames an existing sheet to a new destination name. */
  record RenameSheet(String sheetName, String newSheetName) implements WorkbookCommand {
    public RenameSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(newSheetName, "newSheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (newSheetName.isBlank()) {
        throw new IllegalArgumentException("newSheetName must not be blank");
      }
    }
  }

  /** Deletes an existing sheet from the workbook. */
  record DeleteSheet(String sheetName) implements WorkbookCommand {
    public DeleteSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Moves an existing sheet to a zero-based workbook position. */
  record MoveSheet(String sheetName, int targetIndex) implements WorkbookCommand {
    public MoveSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (targetIndex < 0) {
        throw new IllegalArgumentException("targetIndex must not be negative");
      }
    }
  }

  /** Copies one sheet into a new visible, unselected sheet at the requested workbook position. */
  record CopySheet(String sourceSheetName, String newSheetName, ExcelSheetCopyPosition position)
      implements WorkbookCommand {
    public CopySheet {
      Objects.requireNonNull(sourceSheetName, "sourceSheetName must not be null");
      Objects.requireNonNull(newSheetName, "newSheetName must not be null");
      Objects.requireNonNull(position, "position must not be null");
      if (sourceSheetName.isBlank()) {
        throw new IllegalArgumentException("sourceSheetName must not be blank");
      }
      if (newSheetName.isBlank()) {
        throw new IllegalArgumentException("newSheetName must not be blank");
      }
    }
  }

  /** Sets the active sheet and ensures it is selected. */
  record SetActiveSheet(String sheetName) implements WorkbookCommand {
    public SetActiveSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Sets the selected visible sheet set. */
  record SetSelectedSheets(List<String> sheetNames) implements WorkbookCommand {
    public SetSelectedSheets {
      Objects.requireNonNull(sheetNames, "sheetNames must not be null");
      sheetNames = List.copyOf(sheetNames);
      if (sheetNames.isEmpty()) {
        throw new IllegalArgumentException("sheetNames must not be empty");
      }
      for (String sheetName : sheetNames) {
        Objects.requireNonNull(sheetName, "sheetNames must not contain nulls");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetNames must not contain blank values");
        }
      }
      if (sheetNames.size() != new java.util.LinkedHashSet<>(sheetNames).size()) {
        throw new IllegalArgumentException("sheetNames must not contain duplicates");
      }
    }
  }

  /** Sets one sheet visibility. */
  record SetSheetVisibility(String sheetName, ExcelSheetVisibility visibility)
      implements WorkbookCommand {
    public SetSheetVisibility {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(visibility, "visibility must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Enables sheet protection with the exact supported lock flags. */
  record SetSheetProtection(String sheetName, ExcelSheetProtectionSettings protection)
      implements WorkbookCommand {
    public SetSheetProtection {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Disables sheet protection entirely. */
  record ClearSheetProtection(String sheetName) implements WorkbookCommand {
    public ClearSheetProtection {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Merges an A1-style rectangular range into one displayed cell region. */
  record MergeCells(String sheetName, String range) implements WorkbookCommand {
    public MergeCells {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Removes the merged region whose coordinates exactly match the given range. */
  record UnmergeCells(String sheetName, String range) implements WorkbookCommand {
    public UnmergeCells {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Sets the width of one or more contiguous columns in Excel character units. */
  record SetColumnWidth(
      String sheetName, int firstColumnIndex, int lastColumnIndex, double widthCharacters)
      implements WorkbookCommand {
    public SetColumnWidth {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (firstColumnIndex < 0) {
        throw new IllegalArgumentException("firstColumnIndex must not be negative");
      }
      if (lastColumnIndex < 0) {
        throw new IllegalArgumentException("lastColumnIndex must not be negative");
      }
      if (lastColumnIndex < firstColumnIndex) {
        throw new IllegalArgumentException(
            "lastColumnIndex must not be less than firstColumnIndex");
      }
      requireColumnWidthCharacters(widthCharacters);
    }
  }

  /** Sets the height of one or more contiguous rows in Excel point units. */
  record SetRowHeight(String sheetName, int firstRowIndex, int lastRowIndex, double heightPoints)
      implements WorkbookCommand {
    public SetRowHeight {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (firstRowIndex < 0) {
        throw new IllegalArgumentException("firstRowIndex must not be negative");
      }
      if (lastRowIndex < 0) {
        throw new IllegalArgumentException("lastRowIndex must not be negative");
      }
      if (lastRowIndex < firstRowIndex) {
        throw new IllegalArgumentException("lastRowIndex must not be less than firstRowIndex");
      }
      requireRowHeightPoints(heightPoints);
    }
  }

  /** Applies one explicit pane state to a sheet. */
  record SetSheetPane(String sheetName, ExcelSheetPane pane) implements WorkbookCommand {
    public SetSheetPane {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(pane, "pane must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Applies one explicit zoom percentage to a sheet. */
  record SetSheetZoom(String sheetName, int zoomPercent) implements WorkbookCommand {
    public SetSheetZoom {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      ExcelSheetViewSupport.requireZoomPercent(zoomPercent);
    }
  }

  /** Applies one authoritative supported print-layout state to a sheet. */
  record SetPrintLayout(String sheetName, ExcelPrintLayout printLayout) implements WorkbookCommand {
    public SetPrintLayout {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(printLayout, "printLayout must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Clears the supported print-layout state from a sheet. */
  record ClearPrintLayout(String sheetName) implements WorkbookCommand {
    public ClearPrintLayout {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  record SetCell(String sheetName, String address, ExcelCellValue value)
      implements WorkbookCommand {
    public SetCell {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(value, "value must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  record SetRange(String sheetName, String range, List<List<ExcelCellValue>> rows)
      implements WorkbookCommand {
    public SetRange {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
      rows = List.copyOf(rows);
      if (rows.isEmpty()) {
        throw new IllegalArgumentException("rows must not be empty");
      }
      int expectedWidth = -1;
      for (List<ExcelCellValue> row : rows) {
        Objects.requireNonNull(row, "rows must not contain nulls");
        List<ExcelCellValue> copiedRow = List.copyOf(row);
        if (copiedRow.isEmpty()) {
          throw new IllegalArgumentException("rows must not contain empty rows");
        }
        if (expectedWidth < 0) {
          expectedWidth = copiedRow.size();
        } else if (copiedRow.size() != expectedWidth) {
          throw new IllegalArgumentException("rows must describe a rectangular matrix");
        }
        for (ExcelCellValue value : copiedRow) {
          Objects.requireNonNull(value, "rows must not contain null cell values");
        }
      }
    }
  }

  record ClearRange(String sheetName, String range) implements WorkbookCommand {
    public ClearRange {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Replaces the hyperlink attached to a single cell. */
  record SetHyperlink(String sheetName, String address, ExcelHyperlink target)
      implements WorkbookCommand {
    public SetHyperlink {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(target, "target must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Removes any hyperlink attached to a single existing cell. */
  record ClearHyperlink(String sheetName, String address) implements WorkbookCommand {
    public ClearHyperlink {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Replaces the plain-text comment attached to a single cell. */
  record SetComment(String sheetName, String address, ExcelComment comment)
      implements WorkbookCommand {
    public SetComment {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(comment, "comment must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  /** Removes any comment attached to a single existing cell. */
  record ClearComment(String sheetName, String address) implements WorkbookCommand {
    public ClearComment {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  record ApplyStyle(String sheetName, String range, ExcelCellStyle style)
      implements WorkbookCommand {
    public ApplyStyle {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(style, "style must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  record SetDataValidation(String sheetName, String range, ExcelDataValidationDefinition validation)
      implements WorkbookCommand {
    public SetDataValidation {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(validation, "validation must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  record ClearDataValidations(String sheetName, ExcelRangeSelection selection)
      implements WorkbookCommand {
    public ClearDataValidations {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one logical conditional-formatting block over the requested ranges. */
  record SetConditionalFormatting(String sheetName, ExcelConditionalFormattingBlockDefinition block)
      implements WorkbookCommand {
    public SetConditionalFormatting {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(block, "block must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Removes conditional-formatting blocks on the sheet that intersect the provided selection. */
  record ClearConditionalFormatting(String sheetName, ExcelRangeSelection selection)
      implements WorkbookCommand {
    public ClearConditionalFormatting {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(selection, "selection must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one sheet-level autofilter range. */
  record SetAutofilter(String sheetName, String range) implements WorkbookCommand {
    public SetAutofilter {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  record ClearAutofilter(String sheetName) implements WorkbookCommand {
    public ClearAutofilter {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one workbook-global table definition. */
  record SetTable(ExcelTableDefinition definition) implements WorkbookCommand {
    public SetTable {
      Objects.requireNonNull(definition, "definition must not be null");
    }
  }

  /** Deletes one existing table by workbook-global name and expected sheet. */
  record DeleteTable(String name, String sheetName) implements WorkbookCommand {
    public DeleteTable {
      name = ExcelTableDefinition.validateName(name);
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one named range in workbook or sheet scope. */
  record SetNamedRange(ExcelNamedRangeDefinition definition) implements WorkbookCommand {
    public SetNamedRange {
      Objects.requireNonNull(definition, "definition must not be null");
    }
  }

  /** Deletes one existing named range from workbook or sheet scope. */
  record DeleteNamedRange(String name, ExcelNamedRangeScope scope) implements WorkbookCommand {
    public DeleteNamedRange {
      name = ExcelNamedRangeDefinition.validateName(name);
      Objects.requireNonNull(scope, "scope must not be null");
    }
  }

  record AppendRow(String sheetName, List<ExcelCellValue> values) implements WorkbookCommand {
    public AppendRow {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(values, "values must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      values = List.copyOf(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
      for (ExcelCellValue value : values) {
        Objects.requireNonNull(value, "value must not be null");
      }
    }
  }

  /** Auto-sizes all populated columns on the named sheet. */
  record AutoSizeColumns(String sheetName) implements WorkbookCommand {
    public AutoSizeColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  record EvaluateAllFormulas() implements WorkbookCommand {}

  record ForceFormulaRecalculationOnOpen() implements WorkbookCommand {}

  private static void requireColumnWidthCharacters(double widthCharacters) {
    requireFinitePositive(widthCharacters, "widthCharacters");
    if (widthCharacters > 255.0d) {
      throw new IllegalArgumentException(
          "widthCharacters must be less than or equal to 255.0: " + widthCharacters);
    }
    if (Math.round(widthCharacters * 256.0d) <= 0) {
      throw new IllegalArgumentException(
          "widthCharacters is too small to produce a visible Excel column width: "
              + widthCharacters);
    }
  }

  private static void requireRowHeightPoints(double heightPoints) {
    requireFinitePositive(heightPoints, "heightPoints");
    if (heightPoints > Short.MAX_VALUE / 20.0d) {
      throw new IllegalArgumentException(
          "heightPoints is too large for Excel row height storage: " + heightPoints);
    }
    if ((long) (heightPoints * 20.0d) <= 0) {
      throw new IllegalArgumentException(
          "heightPoints is too small to produce a visible Excel row height: " + heightPoints);
    }
  }

  private static void requireFinitePositive(double value, String fieldName) {
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    if (value <= 0.0d) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
  }
}
