package dev.erst.gridgrind.protocol.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.protocol.catalog.gather.CatalogDuplicateFailures;
import dev.erst.gridgrind.protocol.catalog.gather.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.protocol.catalog.gather.CatalogGatherers;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Stream;

/** Publishes the machine-readable GridGrind protocol surface used by CLI discovery commands. */
public final class GridGrindProtocolCatalog {
  private static final String DISCRIMINATOR_FIELD = "type";
  private static final GridGrindRequest REQUEST_TEMPLATE =
      new GridGrindRequest(
          GridGrindProtocolVersion.current(),
          new GridGrindRequest.WorkbookSource.New(),
          new GridGrindRequest.WorkbookPersistence.None(),
          List.of(),
          List.of());
  private static final List<TypeDescriptor> SOURCE_TYPES =
      List.of(
          descriptor(
              GridGrindRequest.WorkbookSource.New.class,
              "NEW",
              "Create a brand-new empty workbook. A new workbook starts with zero sheets;"
                  + " use ENSURE_SHEET to create the first sheet."),
          descriptor(
              GridGrindRequest.WorkbookSource.ExistingFile.class,
              "EXISTING",
              "Open an existing .xlsx workbook from disk."
                  + " Relative paths resolve in the current execution environment."));
  private static final List<TypeDescriptor> PERSISTENCE_TYPES =
      List.of(
          descriptor(
              GridGrindRequest.WorkbookPersistence.None.class,
              "NONE",
              "Keep the workbook in memory only." + " The response persistence.type echoes NONE."),
          descriptor(
              GridGrindRequest.WorkbookPersistence.OverwriteSource.class,
              "OVERWRITE",
              "Overwrite the opened source workbook at source.path."
                  + " No path field is accepted on OVERWRITE;"
                  + " the write target is the same path opened by the EXISTING source."
                  + " Relative source.path values resolve in the current execution environment."
                  + " The response persistence.type echoes OVERWRITE and includes sourcePath"
                  + " (the original source path string) and executionPath (absolute normalized)."),
          descriptor(
              GridGrindRequest.WorkbookPersistence.SaveAs.class,
              "SAVE_AS",
              "Save the workbook to a new .xlsx path."
                  + " Relative paths resolve in the current execution environment."
                  + " The response persistence.type echoes SAVE_AS and includes requestedPath"
                  + " (the literal path from the request) and executionPath (the absolute"
                  + " normalized path where the file was written); they differ when a relative"
                  + " path or a path with .. segments is supplied."
                  + " Missing parent directories are created automatically."));
  private static final List<TypeDescriptor> OPERATION_TYPES =
      List.of(
          descriptor(
              WorkbookOperation.EnsureSheet.class,
              "ENSURE_SHEET",
              "Create the sheet if it does not already exist."
                  + " Sheet names must not exceed 31 characters (Excel limit)."),
          descriptor(
              WorkbookOperation.RenameSheet.class,
              "RENAME_SHEET",
              "Rename an existing sheet."
                  + " The new name must not exceed 31 characters (Excel limit)."),
          descriptor(
              WorkbookOperation.DeleteSheet.class,
              "DELETE_SHEET",
              "Delete an existing sheet."
                  + " A workbook must retain at least one sheet and at least one visible sheet;"
                  + " deleting the last sheet or the last visible sheet returns INVALID_REQUEST."),
          descriptor(
              WorkbookOperation.MoveSheet.class,
              "MOVE_SHEET",
              "Move a sheet to a zero-based workbook position."
                  + " targetIndex is 0-based: 0 moves the sheet to the front,"
                  + " sheetCount-1 moves it to the back."),
          descriptor(
              WorkbookOperation.CopySheet.class,
              "COPY_SHEET",
              "Copy one sheet into a new visible, unselected sheet."
                  + " position defaults to APPEND_AT_END when omitted."
                  + " The copied sheet preserves supported sheet-local workbook content such as"
                  + " formulas, validations, conditional formatting, comments, hyperlinks,"
                  + " merged regions, and layout state."
                  + " Sheets containing tables, or sheet-scoped formula-defined named ranges,"
                  + " are rejected explicitly because they are not copyable under the current"
                  + " product contract.",
              "position"),
          descriptor(
              WorkbookOperation.SetActiveSheet.class,
              "SET_ACTIVE_SHEET",
              "Set the active sheet."
                  + " Hidden sheets cannot be activated."
                  + " The active sheet is always selected."),
          descriptor(
              WorkbookOperation.SetSelectedSheets.class,
              "SET_SELECTED_SHEETS",
              "Set the selected visible sheet set."
                  + " Duplicate or unknown sheet names are rejected."
                  + " selectedSheetNames read back in workbook order;"
                  + " activeSheetName preserves the primary selected sheet choice."),
          descriptor(
              WorkbookOperation.SetSheetVisibility.class,
              "SET_SHEET_VISIBILITY",
              "Set one sheet visibility state."
                  + " A workbook must retain at least one visible sheet;"
                  + " hiding the last visible sheet is rejected."),
          descriptor(
              WorkbookOperation.SetSheetProtection.class,
              "SET_SHEET_PROTECTION",
              "Enable sheet protection with the exact supported lock flags."
                  + " Password-bearing protection is intentionally out of scope."),
          descriptor(
              WorkbookOperation.ClearSheetProtection.class,
              "CLEAR_SHEET_PROTECTION",
              "Disable sheet protection entirely."),
          descriptor(
              WorkbookOperation.MergeCells.class,
              "MERGE_CELLS",
              "Merge a rectangular A1-style range."),
          descriptor(
              WorkbookOperation.UnmergeCells.class,
              "UNMERGE_CELLS",
              "Remove one merged region by exact range match."),
          descriptor(
              WorkbookOperation.SetColumnWidth.class,
              "SET_COLUMN_WIDTH",
              "Set one or more column widths in Excel character units."
                  + " widthCharacters must be > 0 and <= 255 (Excel column width limit)."),
          descriptor(
              WorkbookOperation.SetRowHeight.class,
              "SET_ROW_HEIGHT",
              "Set one or more row heights in Excel point units."
                  + " heightPoints must be > 0 and <= 1638.35 (Excel row height limit: 32767 twips)."),
          descriptor(
              WorkbookOperation.SetSheetPane.class,
              "SET_SHEET_PANE",
              "Apply one explicit pane state."
                  + " pane.type can be NONE, FROZEN, or SPLIT; use NONE to clear panes."),
          descriptor(
              WorkbookOperation.SetSheetZoom.class,
              "SET_SHEET_ZOOM",
              "Set the sheet zoom percentage."
                  + " zoomPercent must be between 10 and 400 inclusive."),
          descriptor(
              WorkbookOperation.SetPrintLayout.class,
              "SET_PRINT_LAYOUT",
              "Apply one authoritative supported print-layout state to a sheet."
                  + " Omitted nested fields normalize to default or clear state."
                  + " The supported surface covers print area, orientation, fit scaling,"
                  + " repeating rows, repeating columns, and plain header or footer text."),
          descriptor(
              WorkbookOperation.ClearPrintLayout.class,
              "CLEAR_PRINT_LAYOUT",
              "Clear the supported print-layout state from a sheet."),
          descriptor(
              WorkbookOperation.SetCell.class,
              "SET_CELL",
              "Write one typed value to a single cell."),
          descriptor(
              WorkbookOperation.SetRange.class,
              "SET_RANGE",
              "Write a rectangular grid of typed values."),
          descriptor(
              WorkbookOperation.ClearRange.class,
              "CLEAR_RANGE",
              "Clear value, style, hyperlink, and comment state from a range."),
          descriptor(
              WorkbookOperation.SetHyperlink.class,
              "SET_HYPERLINK",
              "Attach a hyperlink to one cell."
                  + " FILE targets use the field name path and normalize file: URIs to plain"
                  + " path strings."),
          descriptor(
              WorkbookOperation.ClearHyperlink.class,
              "CLEAR_HYPERLINK",
              "Remove the hyperlink from one cell; no-op when the cell does not physically exist."),
          descriptor(
              WorkbookOperation.SetComment.class,
              "SET_COMMENT",
              "Attach a plain-text comment to one cell."),
          descriptor(
              WorkbookOperation.ClearComment.class,
              "CLEAR_COMMENT",
              "Remove the comment from one cell; no-op when the cell does not physically exist."),
          descriptor(
              WorkbookOperation.ApplyStyle.class,
              "APPLY_STYLE",
              "Apply a style patch to every cell in a range."
                  + " Write shape: style.border is a nested object"
                  + " { \"all\": { \"style\": \"THIN\" } } or per-side top/right/bottom/left."
                  + " Read shape (GET_CELLS, GET_WINDOW): borders are flat properties"
                  + " topBorderStyle, rightBorderStyle, bottomBorderStyle, leftBorderStyle;"
                  + " the nested border object is write-only."),
          descriptor(
              WorkbookOperation.SetDataValidation.class,
              "SET_DATA_VALIDATION",
              "Create or replace one data-validation rule over the supplied sheet range."
                  + " Overlapping existing rules are normalized so the written rule becomes"
                  + " authoritative on its target range."),
          descriptor(
              WorkbookOperation.ClearDataValidations.class,
              "CLEAR_DATA_VALIDATIONS",
              "Remove data-validation structures from the selected ranges on one sheet."
                  + " SELECTED removes only intersecting coverage; ALL clears every rule"
                  + " on the sheet."),
          descriptor(
              WorkbookOperation.SetConditionalFormatting.class,
              "SET_CONDITIONAL_FORMATTING",
              "Create or replace one logical conditional-formatting block over the supplied"
                  + " sheet ranges."
                  + " The write contract currently authors formula rules and cell-value rules."
                  + " Any existing conditional-formatting block that intersects the target ranges"
                  + " is removed first so the written block becomes authoritative on that"
                  + " coverage."),
          descriptor(
              WorkbookOperation.ClearConditionalFormatting.class,
              "CLEAR_CONDITIONAL_FORMATTING",
              "Remove conditional-formatting blocks from the selected ranges on one sheet."
                  + " SELECTED removes whole blocks whose stored ranges intersect the supplied"
                  + " ranges; ALL clears every conditional-formatting block on the sheet."),
          descriptor(
              WorkbookOperation.SetAutofilter.class,
              "SET_AUTOFILTER",
              "Create or replace one sheet-level autofilter range."
                  + " The range must include a nonblank header row and must not overlap"
                  + " an existing table range."),
          descriptor(
              WorkbookOperation.ClearAutofilter.class,
              "CLEAR_AUTOFILTER",
              "Clear the sheet-level autofilter range on one sheet."
                  + " Table-owned autofilters remain attached to their tables."),
          descriptor(
              WorkbookOperation.SetTable.class,
              "SET_TABLE",
              "Create or replace one workbook-global table definition."
                  + " Table names are workbook-global and case-insensitive."
                  + " Header cells must be nonblank and unique (case-insensitive)."
                  + " Overlapping existing tables are rejected."
                  + " Any overlapping sheet-level autofilter is cleared so the table-owned"
                  + " autofilter becomes authoritative on that range."),
          descriptor(
              WorkbookOperation.DeleteTable.class,
              "DELETE_TABLE",
              "Delete one existing workbook-global table by name and expected sheet name."),
          descriptor(
              WorkbookOperation.SetNamedRange.class,
              "SET_NAMED_RANGE",
              "Create or replace one workbook- or sheet-scoped named range."),
          descriptor(
              WorkbookOperation.DeleteNamedRange.class,
              "DELETE_NAMED_RANGE",
              "Delete one existing workbook- or sheet-scoped named range."),
          descriptor(
              WorkbookOperation.AppendRow.class,
              "APPEND_ROW",
              "Append a row of typed values after the last value-bearing row."
                  + " Blank rows that carry only style, comment, or hyperlink metadata do not"
                  + " affect the append position."),
          descriptor(
              WorkbookOperation.AutoSizeColumns.class,
              "AUTO_SIZE_COLUMNS",
              "Size columns deterministically from displayed cell content so the resulting"
                  + " widths are stable in headless and container runs."),
          descriptor(
              WorkbookOperation.EvaluateFormulas.class,
              "EVALUATE_FORMULAS",
              "Evaluate workbook formulas immediately."),
          descriptor(
              WorkbookOperation.ForceFormulaRecalculationOnOpen.class,
              "FORCE_FORMULA_RECALCULATION_ON_OPEN",
              "Mark the workbook to recalculate when opened in Excel."));
  private static final List<TypeDescriptor> READ_TYPES =
      List.of(
          descriptor(
              WorkbookReadOperation.GetWorkbookSummary.class,
              "GET_WORKBOOK_SUMMARY",
              "Return workbook-level summary facts including sheet count, sheet names,"
                  + " named range count, and formula recalculation flag."
                  + " Empty workbooks return workbook.kind=EMPTY;"
                  + " non-empty workbooks return workbook.kind=WITH_SHEETS with activeSheetName"
                  + " and selectedSheetNames."),
          descriptor(
              WorkbookReadOperation.GetNamedRanges.class,
              "GET_NAMED_RANGES",
              "Return named ranges matched by the supplied selection."),
          descriptor(
              WorkbookReadOperation.GetSheetSummary.class,
              "GET_SHEET_SUMMARY",
              "Return structural summary facts for one sheet."
                  + " Includes visibility and sheet protection state."
                  + " physicalRowCount is the number of physically materialized rows (sparse)."
                  + " lastRowIndex is the 0-based index of the last materialized row"
                  + " (-1 when empty), including metadata-only rows."
                  + " lastColumnIndex is the 0-based index of the last materialized column"
                  + " in any row (-1 when empty)."),
          descriptor(
              WorkbookReadOperation.GetCells.class,
              "GET_CELLS",
              "Return exact cell snapshots for explicit addresses."
                  + " An invalid or out-of-range address returns INVALID_CELL_ADDRESS, not a blank."
                  + " Each snapshot includes address, declaredType, effectiveType, displayValue,"
                  + " style, and metadata fields. Type-specific fields: stringValue (STRING),"
                  + " numberValue (NUMBER), booleanValue (BOOLEAN), errorValue (ERROR),"
                  + " formula and evaluation (FORMULA). For FORMULA cells, effectiveType is FORMULA"
                  + " and the evaluated result type is in evaluation.effectiveType."
                  + " style.fontHeight is a plain object with both twips and points fields,"
                  + " not the discriminated FontHeightInput write format."),
          descriptor(
              WorkbookReadOperation.GetWindow.class,
              "GET_WINDOW",
              "Return a rectangular window of cell snapshots."
                  + " rowCount * columnCount must not exceed "
                  + WorkbookReadOperation.MAX_WINDOW_CELLS
                  + ". Each cell snapshot has the same shape as GET_CELLS: address,"
                  + " declaredType, effectiveType, displayValue, style, metadata,"
                  + " and type-specific value fields."
                  + " For FORMULA cells, effectiveType is FORMULA and the evaluated result type"
                  + " is in evaluation.effectiveType."
                  + " Response shape: { \"window\": { \"rows\": [ { \"cells\": [...] } ] } }."
                  + " Note: the top-level key is \"window\" and cells are nested under"
                  + " window.rows[N].cells, unlike GET_CELLS which places cells directly"
                  + " under the top-level \"cells\" key."),
          descriptor(
              WorkbookReadOperation.GetMergedRegions.class,
              "GET_MERGED_REGIONS",
              "Return the merged regions defined on one sheet."),
          descriptor(
              WorkbookReadOperation.GetHyperlinks.class,
              "GET_HYPERLINKS",
              "Return hyperlink metadata for selected cells in the same discriminated shape used"
                  + " by SET_HYPERLINK targets. FILE targets are returned as normalized path"
                  + " strings, not file: URIs."),
          descriptor(
              WorkbookReadOperation.GetComments.class,
              "GET_COMMENTS",
              "Return comment metadata for selected cells."),
          descriptor(
              WorkbookReadOperation.GetSheetLayout.class,
              "GET_SHEET_LAYOUT",
              "Return pane, zoom, row-height, and column-width metadata."),
          descriptor(
              WorkbookReadOperation.GetPrintLayout.class,
              "GET_PRINT_LAYOUT",
              "Return supported print-layout metadata for one sheet, including print area,"
                  + " orientation, scaling, repeating rows or columns, and plain header or"
                  + " footer text."),
          descriptor(
              WorkbookReadOperation.GetDataValidations.class,
              "GET_DATA_VALIDATIONS",
              "Return factual data-validation structures for the selected sheet ranges."
                  + " Supported rules include explicit lists, formula lists, comparison rules,"
                  + " and custom formulas; unsupported rules are surfaced explicitly with typed"
                  + " detail."),
          descriptor(
              WorkbookReadOperation.GetConditionalFormatting.class,
              "GET_CONDITIONAL_FORMATTING",
              "Return factual conditional-formatting blocks for the selected sheet ranges."
                  + " Read families include authored formula and cell-value rules plus loaded"
                  + " color scales, data bars, icon sets, and explicitly unsupported rules."
                  + " Each block preserves its stored ordered ranges and rule priority data."),
          descriptor(
              WorkbookReadOperation.GetAutofilters.class,
              "GET_AUTOFILTERS",
              "Return sheet- and table-owned autofilter metadata for one sheet."
                  + " Each entry is typed as SHEET or TABLE so ownership is explicit."),
          descriptor(
              WorkbookReadOperation.GetTables.class,
              "GET_TABLES",
              "Return factual table metadata selected by workbook-global table name or ALL."
                  + " Each table includes range, header and totals row counts, column names,"
                  + " style metadata, and whether a table-owned autofilter is present."),
          descriptor(
              WorkbookReadOperation.GetFormulaSurface.class,
              "GET_FORMULA_SURFACE",
              "Summarize formula usage patterns across the selected sheets."),
          descriptor(
              WorkbookReadOperation.GetSheetSchema.class,
              "GET_SHEET_SCHEMA",
              "Infer a simple schema from a rectangular sheet window."
                  + " rowCount * columnCount must not exceed "
                  + WorkbookReadOperation.MAX_WINDOW_CELLS
                  + "."
                  + " The first row is treated as the header; dataRowCount is 0 when all header"
                  + " cells are blank."
                  + " dominantType is null when all data cells are blank, or when two or more types"
                  + " tie for the highest count."
                  + " Formula cells contribute their evaluated result type (NUMBER, STRING, etc.)"
                  + " to dominantType and observedTypes, not FORMULA."),
          descriptor(
              WorkbookReadOperation.GetNamedRangeSurface.class,
              "GET_NAMED_RANGE_SURFACE",
              "Summarize the scope and backing kind of named ranges."),
          descriptor(
              WorkbookReadOperation.AnalyzeFormulaHealth.class,
              "ANALYZE_FORMULA_HEALTH",
              "Report formula findings such as errors and volatile usage."),
          descriptor(
              WorkbookReadOperation.AnalyzeDataValidationHealth.class,
              "ANALYZE_DATA_VALIDATION_HEALTH",
              "Report data-validation findings such as unsupported, overlapping, or"
                  + " broken-formula rules."),
          descriptor(
              WorkbookReadOperation.AnalyzeConditionalFormattingHealth.class,
              "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
              "Report conditional-formatting findings such as broken formulas,"
                  + " unsupported loaded rules, invalid target ranges, or priority collisions."),
          descriptor(
              WorkbookReadOperation.AnalyzeAutofilterHealth.class,
              "ANALYZE_AUTOFILTER_HEALTH",
              "Report autofilter findings such as invalid ranges, blank header rows,"
                  + " or ownership mismatches between sheet-level filters and tables."),
          descriptor(
              WorkbookReadOperation.AnalyzeTableHealth.class,
              "ANALYZE_TABLE_HEALTH",
              "Report table findings such as overlaps, broken references,"
                  + " blank or duplicate headers, and unresolved styles."),
          descriptor(
              WorkbookReadOperation.AnalyzeHyperlinkHealth.class,
              "ANALYZE_HYPERLINK_HEALTH",
              "Report hyperlink findings such as malformed, missing, unresolved, or broken"
                  + " targets."
                  + " Relative FILE targets are resolved against the workbook's persisted path"
                  + " when one exists."),
          descriptor(
              WorkbookReadOperation.AnalyzeNamedRangeHealth.class,
              "ANALYZE_NAMED_RANGE_HEALTH",
              "Report named-range findings such as broken references."),
          descriptor(
              WorkbookReadOperation.AnalyzeWorkbookFindings.class,
              "ANALYZE_WORKBOOK_FINDINGS",
              "Run all analysis families (formula health, data-validation health,"
                  + " conditional-formatting health, autofilter health, table health,"
                  + " hyperlink health, and named-range health) across the entire workbook"
                  + " and aggregate findings in a single response."));
  private static final List<NestedTypeDescriptor> NESTED_TYPE_GROUPS =
      List.of(
          nestedTypeGroup(
              "cellInputTypes",
              CellInput.class,
              List.of(
                  descriptor(CellInput.Blank.class, "BLANK", "Write an empty cell."),
                  descriptor(CellInput.Text.class, "TEXT", "Write a string cell value."),
                  descriptor(CellInput.Numeric.class, "NUMBER", "Write a numeric cell value."),
                  descriptor(
                      CellInput.BooleanValue.class, "BOOLEAN", "Write a boolean cell value."),
                  descriptor(
                      CellInput.Date.class,
                      "DATE",
                      "Write an ISO-8601 date value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.DateTime.class,
                      "DATE_TIME",
                      "Write an ISO-8601 date-time value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.Formula.class,
                      "FORMULA",
                      "Write an Excel formula. A leading = sign is accepted and stripped automatically; the engine stores the formula without it."))),
          nestedTypeGroup(
              "hyperlinkTargetTypes",
              HyperlinkTarget.class,
              List.of(
                  descriptor(HyperlinkTarget.Url.class, "URL", "Attach an absolute URL target."),
                  descriptor(
                      HyperlinkTarget.Email.class,
                      "EMAIL",
                      "Attach an email target without the mailto: prefix."),
                  descriptor(
                      HyperlinkTarget.File.class,
                      "FILE",
                      "Attach a local or shared file path."
                          + " Accepts plain paths or file: URIs and normalizes to a plain path"
                          + " string."),
                  descriptor(
                      HyperlinkTarget.Document.class,
                      "DOCUMENT",
                      "Attach an internal workbook target."))),
          nestedTypeGroup(
              "paneTypes",
              PaneInput.class,
              List.of(
                  descriptor(
                      PaneInput.None.class, "NONE", "Sheet has no active pane split or freeze."),
                  descriptor(
                      PaneInput.Frozen.class,
                      "FROZEN",
                      "Freeze the sheet at the provided split and visible-origin coordinates."),
                  descriptor(
                      PaneInput.Split.class,
                      "SPLIT",
                      "Apply split panes with explicit split offsets, visible origin, and active pane."))),
          nestedTypeGroup(
              "sheetCopyPositionTypes",
              SheetCopyPosition.class,
              List.of(
                  descriptor(
                      SheetCopyPosition.AppendAtEnd.class,
                      "APPEND_AT_END",
                      "Place the copied sheet after every existing sheet."),
                  descriptor(
                      SheetCopyPosition.AtIndex.class,
                      "AT_INDEX",
                      "Place the copied sheet at the requested zero-based workbook position."))),
          nestedTypeGroup(
              "cellSelectionTypes",
              CellSelection.class,
              List.of(
                  descriptor(
                      CellSelection.AllUsedCells.class,
                      "ALL_USED_CELLS",
                      "Select every physically present cell on the sheet."),
                  descriptor(
                      CellSelection.Selected.class,
                      "SELECTED",
                      "Select only the supplied ordered cell addresses."))),
          nestedTypeGroup(
              "rangeSelectionTypes",
              RangeSelection.class,
              List.of(
                  descriptor(RangeSelection.All.class, "ALL", "Select every matching range."),
                  descriptor(
                      RangeSelection.Selected.class,
                      "SELECTED",
                      "Select only structures whose stored ranges intersect the supplied"
                          + " ordered A1-style ranges."))),
          nestedTypeGroup(
              "sheetSelectionTypes",
              SheetSelection.class,
              List.of(
                  descriptor(
                      SheetSelection.All.class, "ALL", "Select every sheet in workbook order."),
                  descriptor(
                      SheetSelection.Selected.class,
                      "SELECTED",
                      "Select only the supplied ordered sheet names."))),
          nestedTypeGroup(
              "tableSelectionTypes",
              TableSelection.class,
              List.of(
                  descriptor(
                      TableSelection.All.class, "ALL", "Select every table in workbook order."),
                  descriptor(
                      TableSelection.ByNames.class,
                      "BY_NAMES",
                      "Select only the supplied workbook-global table names."))),
          nestedTypeGroup(
              "namedRangeSelectionTypes",
              NamedRangeSelection.class,
              List.of(
                  descriptor(NamedRangeSelection.All.class, "ALL", "Select every named range."),
                  descriptor(
                      NamedRangeSelection.Selected.class,
                      "SELECTED",
                      "Select only the supplied named-range selectors."))),
          nestedTypeGroup(
              "namedRangeScopeTypes",
              NamedRangeScope.class,
              List.of(
                  descriptor(NamedRangeScope.Workbook.class, "WORKBOOK", "Target workbook scope."),
                  descriptor(
                      NamedRangeScope.Sheet.class, "SHEET", "Target one specific sheet scope."))),
          nestedTypeGroup(
              "namedRangeSelectorTypes",
              NamedRangeSelector.class,
              List.of(
                  descriptor(
                      NamedRangeSelector.ByName.class,
                      "BY_NAME",
                      "Match a named range across all scopes by exact name."),
                  descriptor(
                      NamedRangeSelector.WorkbookScope.class,
                      "WORKBOOK_SCOPE",
                      "Match the workbook-scoped named range with the exact name."),
                  descriptor(
                      NamedRangeSelector.SheetScope.class,
                      "SHEET_SCOPE",
                      "Match the sheet-scoped named range on one sheet."))),
          nestedTypeGroup(
              "fontHeightTypes",
              FontHeightInput.class,
              List.of(
                  descriptor(
                      FontHeightInput.Points.class,
                      "POINTS",
                      "Specify font height in point units."
                          + " Write format: {\"type\":\"POINTS\",\"points\":13}."
                          + " Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is"
                          + " {\"twips\":260,\"points\":13} with both fields present,"
                          + " not this discriminated type format."),
                  descriptor(
                      FontHeightInput.Twips.class,
                      "TWIPS",
                      "Specify font height in exact twips (20 twips = 1 point)."
                          + " Write format: {\"type\":\"TWIPS\",\"twips\":260}."
                          + " Read-back returns the same plain object shape as POINTS."))),
          nestedTypeGroup(
              "dataValidationRuleTypes",
              DataValidationRuleInput.class,
              List.of(
                  descriptor(
                      DataValidationRuleInput.ExplicitList.class,
                      "EXPLICIT_LIST",
                      "Allow only one of the supplied explicit values."),
                  descriptor(
                      DataValidationRuleInput.FormulaList.class,
                      "FORMULA_LIST",
                      "Allow values from a formula-driven list expression."),
                  descriptor(
                      DataValidationRuleInput.WholeNumber.class,
                      "WHOLE_NUMBER",
                      "Apply a whole-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DecimalNumber.class,
                      "DECIMAL_NUMBER",
                      "Apply a decimal-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DateRule.class,
                      "DATE",
                      "Apply a date comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TimeRule.class,
                      "TIME",
                      "Apply a time comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TextLength.class,
                      "TEXT_LENGTH",
                      "Apply a text-length comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.CustomFormula.class,
                      "CUSTOM_FORMULA",
                      "Allow values that satisfy a custom formula."))),
          nestedTypeGroup(
              "conditionalFormattingRuleTypes",
              ConditionalFormattingRuleInput.class,
              List.of(
                  descriptor(
                      ConditionalFormattingRuleInput.FormulaRule.class,
                      "FORMULA_RULE",
                      "Apply one formula-driven conditional-formatting rule."
                          + " Supply one differential style."),
                  descriptor(
                      ConditionalFormattingRuleInput.CellValueRule.class,
                      "CELL_VALUE_RULE",
                      "Apply one cell-value comparison conditional-formatting rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN."
                          + " Supply one differential style.",
                      "formula2"))),
          nestedTypeGroup(
              "printAreaTypes",
              PrintAreaInput.class,
              List.of(
                  descriptor(
                      PrintAreaInput.None.class, "NONE", "Sheet has no explicit print area."),
                  descriptor(
                      PrintAreaInput.Range.class,
                      "RANGE",
                      "Sheet prints the provided rectangular A1-style range."))),
          nestedTypeGroup(
              "printScalingTypes",
              PrintScalingInput.class,
              List.of(
                  descriptor(
                      PrintScalingInput.Automatic.class,
                      "AUTOMATIC",
                      "Sheet uses Excel's default scaling instead of fit-to-page counts."),
                  descriptor(
                      PrintScalingInput.Fit.class,
                      "FIT",
                      "Sheet fits printed content into the provided page counts."
                          + " A value of 0 on one axis keeps that axis unconstrained."))),
          nestedTypeGroup(
              "printTitleRowsTypes",
              PrintTitleRowsInput.class,
              List.of(
                  descriptor(
                      PrintTitleRowsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title rows."),
                  descriptor(
                      PrintTitleRowsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based row band on every printed page."))),
          nestedTypeGroup(
              "printTitleColumnsTypes",
              PrintTitleColumnsInput.class,
              List.of(
                  descriptor(
                      PrintTitleColumnsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title columns."),
                  descriptor(
                      PrintTitleColumnsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based column band on every printed page."))),
          nestedTypeGroup(
              "tableStyleTypes",
              TableStyleInput.class,
              List.of(
                  descriptor(TableStyleInput.None.class, "NONE", "Clear table style metadata."),
                  descriptor(
                      TableStyleInput.Named.class,
                      "NAMED",
                      "Apply one named workbook table style with explicit stripe and emphasis"
                          + " flags."))));
  private static final List<PlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "commentInputType",
              CommentInput.class,
              "CommentInput",
              "Plain-text comment payload attached to one cell.",
              List.of("visible")),
          plainTypeDescriptor(
              "namedRangeTargetType",
              NamedRangeTarget.class,
              "NamedRangeTarget",
              "Explicit sheet name and A1-style range address for named-range authoring.",
              List.of()),
          plainTypeDescriptor(
              "sheetProtectionSettingsType",
              SheetProtectionSettings.class,
              "SheetProtectionSettings",
              "Supported sheet-protection lock flags authored and reported by GridGrind.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleInputType",
              CellStyleInput.class,
              "CellStyleInput",
              "Style patch applied to a cell or range; at least one field must be set."
                  + " Colors use #RRGGBB hex.",
              List.of(
                  "numberFormat",
                  "bold",
                  "italic",
                  "wrapText",
                  "horizontalAlignment",
                  "verticalAlignment",
                  "fontName",
                  "fontHeight",
                  "fontColor",
                  "underline",
                  "strikeout",
                  "fillColor",
                  "border")),
          plainTypeDescriptor(
              "cellBorderInputType",
              CellBorderInput.class,
              "CellBorderInput",
              "Border patch for cell styling; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "cellBorderSideInputType",
              CellBorderSideInput.class,
              "CellBorderSideInput",
              "One border side defined by its border style.",
              List.of()),
          plainTypeDescriptor(
              "dataValidationInputType",
              DataValidationInput.class,
              "DataValidationInput",
              "Supported data-validation definition attached to one sheet range.",
              List.of("allowBlank", "suppressDropDownArrow", "prompt", "errorAlert")),
          plainTypeDescriptor(
              "dataValidationPromptInputType",
              DataValidationPromptInput.class,
              "DataValidationPromptInput",
              "Optional prompt-box configuration shown when a validated cell is selected.",
              List.of("showPromptBox")),
          plainTypeDescriptor(
              "dataValidationErrorAlertInputType",
              DataValidationErrorAlertInput.class,
              "DataValidationErrorAlertInput",
              "Optional error-box configuration shown when invalid data is entered.",
              List.of("showErrorBox")),
          plainTypeDescriptor(
              "conditionalFormattingBlockInputType",
              ConditionalFormattingBlockInput.class,
              "ConditionalFormattingBlockInput",
              "One authored conditional-formatting block with ordered target ranges and rules."
                  + " rules must not be empty; ranges must be unique.",
              List.of()),
          plainTypeDescriptor(
              "headerFooterTextInputType",
              HeaderFooterTextInput.class,
              "HeaderFooterTextInput",
              "Plain left, center, and right header or footer text segments."
                  + " Null fields default to empty string.",
              List.of("left", "center", "right")),
          plainTypeDescriptor(
              "differentialStyleInputType",
              DifferentialStyleInput.class,
              "DifferentialStyleInput",
              "Differential style payload used by authored conditional-formatting rules."
                  + " At least one field must be set. Colors use #RRGGBB hex.",
              List.of(
                  "numberFormat",
                  "bold",
                  "italic",
                  "fontHeight",
                  "fontColor",
                  "underline",
                  "strikeout",
                  "fillColor",
                  "border")),
          plainTypeDescriptor(
              "differentialBorderInputType",
              DifferentialBorderInput.class,
              "DifferentialBorderInput",
              "Conditional-formatting differential border patch; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "differentialBorderSideInputType",
              DifferentialBorderSideInput.class,
              "DifferentialBorderSideInput",
              "One conditional-formatting differential border side defined by style and optional"
                  + " color.",
              List.of()),
          plainTypeDescriptor(
              "printLayoutInputType",
              PrintLayoutInput.class,
              "PrintLayoutInput",
              "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "printArea",
                  "orientation",
                  "scaling",
                  "repeatingRows",
                  "repeatingColumns",
                  "header",
                  "footer")),
          plainTypeDescriptor(
              "tableInputType",
              TableInput.class,
              "TableInput",
              "Workbook-global table definition for one SET_TABLE request.",
              List.of("showTotalsRow")));
  private static final Catalog CATALOG = buildCatalog();

  private GridGrindProtocolCatalog() {}

  /** Returns the minimal successful request emitted by the CLI template command. */
  public static GridGrindRequest requestTemplate() {
    return REQUEST_TEMPLATE;
  }

  /** Returns the machine-readable protocol catalog emitted by the CLI discovery command. */
  public static Catalog catalog() {
    return CATALOG;
  }

  /**
   * Returns the single catalog entry matching the given operation id, or empty when no entry with
   * that id exists. Searches sourceTypes, persistenceTypes, operationTypes, and readTypes.
   */
  public static Optional<TypeEntry> entryFor(String id) {
    return allEntries().stream().filter(e -> e.id().equals(id)).findFirst();
  }

  private static List<TypeEntry> allEntries() {
    return Stream.concat(
            Stream.of(
                    CATALOG.sourceTypes(),
                    CATALOG.persistenceTypes(),
                    CATALOG.operationTypes(),
                    CATALOG.readTypes())
                .flatMap(List::stream),
            CATALOG.nestedTypes().stream().map(NestedTypeGroup::types).flatMap(List::stream))
        .toList();
  }

  private static Catalog buildCatalog() {
    validateFieldShapeGroupMappings();
    validateCoverage(GridGrindRequest.WorkbookSource.class, SOURCE_TYPES);
    validateCoverage(GridGrindRequest.WorkbookPersistence.class, PERSISTENCE_TYPES);
    validateCoverage(WorkbookOperation.class, OPERATION_TYPES);
    validateCoverage(WorkbookReadOperation.class, READ_TYPES);
    for (NestedTypeDescriptor nestedTypeGroup : NESTED_TYPE_GROUPS) {
      validateCoverage(nestedTypeGroup.sealedType(), nestedTypeGroup.typeDescriptors());
    }
    return new Catalog(
        GridGrindProtocolVersion.current(),
        DISCRIMINATOR_FIELD,
        publicEntries(SOURCE_TYPES),
        publicEntries(PERSISTENCE_TYPES),
        publicEntries(OPERATION_TYPES),
        publicEntries(READ_TYPES),
        NESTED_TYPE_GROUPS.stream().map(GridGrindProtocolCatalog::publicGroup).toList(),
        PLAIN_TYPE_DESCRIPTORS.stream().map(GridGrindProtocolCatalog::publicPlainGroup).toList());
  }

  private static NestedTypeGroup publicGroup(NestedTypeDescriptor descriptor) {
    return new NestedTypeGroup(descriptor.group(), publicEntries(descriptor.typeDescriptors()));
  }

  private static PlainTypeGroup publicPlainGroup(PlainTypeDescriptor descriptor) {
    return new PlainTypeGroup(descriptor.group(), descriptor.typeEntry());
  }

  private static List<TypeEntry> publicEntries(List<TypeDescriptor> descriptors) {
    return descriptors.stream().map(TypeDescriptor::typeEntry).toList();
  }

  private static NestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<TypeDescriptor> typeDescriptors) {
    return new NestedTypeDescriptor(group, sealedType, typeDescriptors);
  }

  private static PlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return new PlainTypeDescriptor(
        group, recordType, typeEntry(recordType, id, summary, optionalFields));
  }

  private static TypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new TypeDescriptor(
        recordType, typeEntry(recordType, id, summary, List.of(optionalFields)));
  }

  private static TypeEntry typeEntry(
      Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
    return new TypeEntry(id, summary, fieldEntries(recordType, optionalFields));
  }

  private static List<FieldEntry> fieldEntries(
      Class<? extends Record> recordType, List<String> optionalFields) {
    requiredFields(recordType, optionalFields);
    Set<String> optionalFieldSet = Set.copyOf(optionalFields);
    return Arrays.stream(recordType.getRecordComponents())
        .gather(CatalogGatherers.expandFieldsWithMetadata(optionalFieldSet))
        .toList();
  }

  private static void validateFieldShapeGroupMappings() {
    Set<Class<?>> descriptorNestedTypes =
        NESTED_TYPE_GROUPS.stream()
            .map(NestedTypeDescriptor::sealedType)
            .collect(java.util.stream.Collectors.toSet());
    Set<Class<?>> descriptorPlainTypes =
        PLAIN_TYPE_DESCRIPTORS.stream()
            .map(PlainTypeDescriptor::recordType)
            .collect(java.util.stream.Collectors.toSet());

    for (NestedTypeDescriptor descriptor : NESTED_TYPE_GROUPS) {
      CatalogFieldMetadataSupport.validateNestedTypeGroupMapping(
          descriptor.sealedType(), descriptor.group());
    }
    for (PlainTypeDescriptor descriptor : PLAIN_TYPE_DESCRIPTORS) {
      CatalogFieldMetadataSupport.validatePlainTypeGroupMapping(
          descriptor.recordType(), descriptor.group());
    }

    // Reverse check: every registered type must appear in a descriptor.
    validateReverseGroupMappings(descriptorNestedTypes, descriptorPlainTypes);
  }

  /**
   * Validates that every type registered in the field-shape maps appears in one of the provided
   * descriptor sets. Exposed as package-private so tests can exercise the failure paths with
   * synthetic descriptor sets that are intentionally incomplete.
   */
  static void validateReverseGroupMappings(
      Set<Class<?>> descriptorNestedTypes, Set<Class<?>> descriptorPlainTypes) {
    for (Class<?> registeredType : CatalogFieldMetadataSupport.registeredNestedTypes()) {
      if (!descriptorNestedTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape nested group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
    for (Class<?> registeredType : CatalogFieldMetadataSupport.registeredPlainTypes()) {
      if (!descriptorPlainTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape plain group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
  }

  /** Returns the required record fields after removing the explicitly optional ones. */
  static List<String> requiredFields(
      Class<? extends Record> recordType, List<String> optionalFields) {
    List<String> recordFields = recordFields(recordType);
    for (String optionalField : optionalFields) {
      if (!recordFields.contains(optionalField)) {
        throw new IllegalStateException(
            "Catalog optional field '%s' does not exist on %s"
                .formatted(optionalField, recordType.getName()));
      }
    }
    return recordFields.stream().filter(field -> !optionalFields.contains(field)).toList();
  }

  private static List<String> recordFields(Class<? extends Record> recordType) {
    return Arrays.stream(recordType.getRecordComponents()).map(RecordComponent::getName).toList();
  }

  private static void validateCoverage(Class<?> sealedType, List<TypeDescriptor> descriptors) {
    validateCoverage(
        sealedType,
        toOrderedMap(
            descriptors,
            TypeDescriptor::recordType,
            descriptor -> descriptor.typeEntry().id(),
            "catalog descriptor"));
  }

  /** Validates that a tagged union and the catalog expose the same ordered {@code type} ids. */
  static void validateCoverage(Class<?> sealedType, Map<Class<?>, String> catalogIds) {
    JsonTypeInfo jsonTypeInfo = sealedType.getAnnotation(JsonTypeInfo.class);
    if (jsonTypeInfo == null || !DISCRIMINATOR_FIELD.equals(jsonTypeInfo.property())) {
      throw new IllegalStateException(
          "Catalog coverage requires %s to use discriminator field '%s'"
              .formatted(sealedType.getName(), DISCRIMINATOR_FIELD));
    }
    JsonSubTypes jsonSubTypes = sealedType.getAnnotation(JsonSubTypes.class);
    if (jsonSubTypes == null) {
      throw new IllegalStateException(
          "Catalog coverage requires @JsonSubTypes on " + sealedType.getName());
    }

    Map<Class<?>, String> annotationIds =
        toOrderedMap(
            Arrays.asList(jsonSubTypes.value()),
            JsonSubTypes.Type::value,
            JsonSubTypes.Type::name,
            "annotation subtype");

    for (Class<?> recordType : catalogIds.keySet()) {
      if (!recordType.isRecord()) {
        throw new IllegalStateException(
            "Catalog entry %s does not target a record type".formatted(recordType));
      }
    }

    if (!annotationIds.keySet().equals(catalogIds.keySet())) {
      throw new IllegalStateException(
          "Catalog coverage mismatch for "
              + sealedType.getName()
              + ": annotated="
              + annotationIds.keySet()
              + ", catalog="
              + catalogIds.keySet());
    }

    for (Map.Entry<Class<?>, String> annotationEntry : annotationIds.entrySet()) {
      String catalogId = catalogIds.get(annotationEntry.getKey());
      if (!annotationEntry.getValue().equals(catalogId)) {
        throw new IllegalStateException(
            "Catalog id mismatch for "
                + annotationEntry.getKey().getName()
                + ": annotation="
                + annotationEntry.getValue()
                + ", catalog="
                + catalogId);
      }
    }
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static <T, K, V> Map<K, V> toOrderedMap(
      List<T> items, Function<T, K> keyFn, Function<T, V> valueFn, String label) {
    Map<K, V> result = new LinkedHashMap<>();
    for (T item : items) {
      K key = keyFn.apply(item);
      V value = valueFn.apply(item);
      if (result.containsKey(key)) {
        throw CatalogDuplicateFailures.duplicateEntryFailure(label, result.get(key), value);
      }
      result.put(key, value);
    }
    return result;
  }

  private record TypeDescriptor(Class<? extends Record> recordType, TypeEntry typeEntry) {
    private TypeDescriptor {
      Objects.requireNonNull(recordType, "recordType must not be null");
      Objects.requireNonNull(typeEntry, "typeEntry must not be null");
    }
  }

  private record PlainTypeDescriptor(
      String group, Class<? extends Record> recordType, TypeEntry typeEntry) {
    private PlainTypeDescriptor {
      group = CatalogRecordValidation.requireNonBlank(group, "group");
      Objects.requireNonNull(recordType, "recordType must not be null");
      Objects.requireNonNull(typeEntry, "typeEntry must not be null");
    }
  }

  private record NestedTypeDescriptor(
      String group, Class<?> sealedType, List<TypeDescriptor> typeDescriptors) {
    private NestedTypeDescriptor {
      group = CatalogRecordValidation.requireNonBlank(group, "group");
      Objects.requireNonNull(sealedType, "sealedType must not be null");
      typeDescriptors = List.copyOf(typeDescriptors);
      for (TypeDescriptor typeDescriptor : typeDescriptors) {
        Objects.requireNonNull(typeDescriptor, "typeDescriptors must not contain nulls");
      }
    }
  }
}
