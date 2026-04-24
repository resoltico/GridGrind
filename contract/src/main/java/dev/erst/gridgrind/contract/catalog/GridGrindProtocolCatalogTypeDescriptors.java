package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.List;

/** Owns the concrete protocol entry descriptors published by the public catalog. */
final class GridGrindProtocolCatalogTypeDescriptors {
  static final List<CatalogTypeDescriptor> STEP_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              MutationStep.class,
              "MUTATION",
              "Execute one mutation action against the selected workbook target."),
          GridGrindProtocolCatalog.descriptor(
              AssertionStep.class,
              "ASSERTION",
              "Verify one authored expectation against the selected workbook target."),
          GridGrindProtocolCatalog.descriptor(
              InspectionStep.class,
              "INSPECTION",
              "Run one factual or analytical inspection query against the selected workbook"
                  + " target."));

  static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookSource.New.class,
              "NEW",
              "Create a brand-new empty workbook. A new workbook starts with zero sheets;"
                  + " use ENSURE_SHEET to create the first sheet."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookSource.ExistingFile.class,
              "EXISTING",
              "Open an existing .xlsx workbook from disk."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " source.security.password unlocks encrypted OOXML packages.",
              "security"));

  static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookPersistence.None.class,
              "NONE",
              "Keep the workbook in memory only." + " The response persistence.type echoes NONE."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookPersistence.OverwriteSource.class,
              "OVERWRITE",
              "Overwrite the opened source workbook at source.path."
                  + " No path field is accepted on OVERWRITE;"
                  + " the write target is the same path opened by the EXISTING source."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes OVERWRITE and includes sourcePath"
                  + " (the original source path string) and executionPath (absolute normalized).",
              "security"),
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookPersistence.SaveAs.class,
              "SAVE_AS",
              "Save the workbook to a new .xlsx path."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes SAVE_AS and includes requestedPath"
                  + " (the literal path from the request) and executionPath (the absolute"
                  + " normalized path where the file was written); they differ when a relative"
                  + " path or a path with .. segments is supplied."
                  + " Missing parent directories are created automatically.",
              "security"));

  static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              MutationAction.EnsureSheet.class,
              "ENSURE_SHEET",
              "Create the sheet if it does not already exist."
                  + " Sheet names must be 1 to 31 characters and must not contain"
                  + " : \\ / ? * [ ] or begin or end with a single quote (Excel limit)."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.RenameSheet.class,
              "RENAME_SHEET",
              "Rename an existing sheet."
                  + " The new name must be 1 to 31 characters and must not contain"
                  + " : \\ / ? * [ ] or begin or end with a single quote (Excel limit)."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.DeleteSheet.class,
              "DELETE_SHEET",
              "Delete an existing sheet."
                  + " A workbook must retain at least one sheet and at least one visible sheet;"
                  + " deleting the last sheet or the last visible sheet returns INVALID_REQUEST."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.MoveSheet.class,
              "MOVE_SHEET",
              "Move a sheet to a zero-based workbook position."
                  + " targetIndex is 0-based: 0 moves the sheet to the front,"
                  + " sheetCount-1 moves it to the back."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.CopySheet.class,
              "COPY_SHEET",
              "Copy one sheet into a new visible, unselected sheet."
                  + " newSheetName follows the same Excel sheet-name rules as ENSURE_SHEET."
                  + " position defaults to APPEND_AT_END when omitted."
                  + " The copied sheet preserves sheet-local workbook content such as formulas,"
                  + " validations, conditional formatting, comments, hyperlinks, merged regions,"
                  + " layout state, sheet-scoped named ranges, sheet autofilters, and tables."
                  + " Copied tables are renamed automatically when needed to keep workbook-global"
                  + " table names unique.",
              "position"),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetActiveSheet.class,
              "SET_ACTIVE_SHEET",
              "Set the active sheet."
                  + " Hidden sheets cannot be activated."
                  + " The active sheet is always selected."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetSelectedSheets.class,
              "SET_SELECTED_SHEETS",
              "Set the selected visible sheet set."
                  + " Duplicate or unknown sheet names are rejected."
                  + " selectedSheetNames read back in workbook order;"
                  + " activeSheetName preserves the primary selected sheet choice."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetSheetVisibility.class,
              "SET_SHEET_VISIBILITY",
              "Set one sheet visibility state."
                  + " A workbook must retain at least one visible sheet;"
                  + " hiding the last visible sheet is rejected."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetSheetProtection.class,
              "SET_SHEET_PROTECTION",
              "Enable sheet protection with the exact supported lock flags."
                  + " password is optional; when provided it is hashed into the sheet-protection"
                  + " metadata.",
              "password"),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearSheetProtection.class,
              "CLEAR_SHEET_PROTECTION",
              "Disable sheet protection entirely."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetWorkbookProtection.class,
              "SET_WORKBOOK_PROTECTION",
              "Enable workbook-level protection and optional workbook or revisions passwords."
                  + " Omitted lock flags normalize to false and omitted passwords are cleared."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearWorkbookProtection.class,
              "CLEAR_WORKBOOK_PROTECTION",
              "Clear workbook-level protection and any stored workbook or revisions passwords."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.MergeCells.class,
              "MERGE_CELLS",
              "Merge a rectangular A1-style range."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.UnmergeCells.class,
              "UNMERGE_CELLS",
              "Remove one merged region by exact range match."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetColumnWidth.class,
              "SET_COLUMN_WIDTH",
              "Set one or more column widths in Excel character units."
                  + " widthCharacters must be > 0 and <= 255 (Excel column width limit)."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetRowHeight.class,
              "SET_ROW_HEIGHT",
              "Set one or more row heights in Excel point units."
                  + " heightPoints must be > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + " (Excel row height limit)."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.InsertRows.class,
              "INSERT_ROWS",
              "Insert one or more blank rows before rowIndex."
                  + " rowIndex must be <= last existing row + 1."
                  + " Append-edge inserts on sparse sheets do not materialize a new physical tail"
                  + " row until content or row metadata exists there."
                  + " GridGrind preserves and retargets existing data validations across the"
                  + " insert, but still rejects row inserts that would move tables or"
                  + " sheet autofilters."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.DeleteRows.class,
              "DELETE_ROWS",
              "Delete one inclusive zero-based row band."
                  + " GridGrind rejects row deletes that would move or truncate tables,"
                  + " sheet autofilters, or data validations, and it also rejects deletes that"
                  + " would truncate range-backed named ranges."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ShiftRows.class,
              "SHIFT_ROWS",
              "Move one inclusive zero-based row band by delta rows."
                  + " delta must not be 0."
                  + " GridGrind rejects row shifts that would move tables, sheet autofilters,"
                  + " or data validations, and it also rejects shifts that would partially move"
                  + " or overwrite range-backed named ranges."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.InsertColumns.class,
              "INSERT_COLUMNS",
              "Insert one or more blank columns before columnIndex."
                  + " columnIndex must be <= last existing column + 1."
                  + " Append-edge inserts on sparse sheets do not materialize a new physical"
                  + " tail column until cells or explicit column metadata exist there."
                  + " GridGrind rejects column inserts when the workbook contains formula cells"
                  + " or formula-defined names, and it still rejects inserts that would move"
                  + " tables or sheet autofilters."
                  + " Existing data validations are preserved and retargeted across the insert."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.DeleteColumns.class,
              "DELETE_COLUMNS",
              "Delete one inclusive zero-based column band."
                  + " GridGrind rejects column deletes when the workbook contains formula cells"
                  + " or formula-defined names, or when the edit would move or truncate tables,"
                  + " sheet autofilters, or data validations, or truncate range-backed named"
                  + " ranges."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ShiftColumns.class,
              "SHIFT_COLUMNS",
              "Move one inclusive zero-based column band by delta columns."
                  + " delta must not be 0."
                  + " GridGrind rejects column shifts when the workbook contains formula cells"
                  + " or formula-defined names, or when the edit would move tables,"
                  + " sheet autofilters, or data validations, or partially move or overwrite"
                  + " range-backed named ranges."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetRowVisibility.class,
              "SET_ROW_VISIBILITY",
              "Set the hidden state for one inclusive zero-based row band."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetColumnVisibility.class,
              "SET_COLUMN_VISIBILITY",
              "Set the hidden state for one inclusive zero-based column band."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.GroupRows.class,
              "GROUP_ROWS",
              "Apply one outline group to an inclusive zero-based row band."
                  + " collapsed defaults to false when omitted."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.UngroupRows.class,
              "UNGROUP_ROWS",
              "Remove outline grouping from one inclusive zero-based row band."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.GroupColumns.class,
              "GROUP_COLUMNS",
              "Apply one outline group to an inclusive zero-based column band."
                  + " collapsed defaults to false when omitted."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.UngroupColumns.class,
              "UNGROUP_COLUMNS",
              "Remove outline grouping from one inclusive zero-based column band."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetSheetPane.class,
              "SET_SHEET_PANE",
              "Apply one explicit pane state."
                  + " pane.type can be NONE, FROZEN, or SPLIT; use NONE to clear panes."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetSheetZoom.class,
              "SET_SHEET_ZOOM",
              "Set the sheet zoom percentage."
                  + " zoomPercent must be between "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MIN_ZOOM_PERCENT
                  + " and "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ZOOM_PERCENT
                  + " inclusive."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetSheetPresentation.class,
              "SET_SHEET_PRESENTATION",
              "Apply one authoritative supported sheet-presentation state to a sheet."
                  + " Omitted nested fields normalize to defaults or clear state."
                  + " The supported surface covers screen display flags, right-to-left mode,"
                  + " tab color, outline summary placement, authored default row and column sizing"
                  + " (defaultColumnWidth > 0 and <= 255; defaultRowHeightPoints > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + "), and ignored-errors ranges."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetPrintLayout.class,
              "SET_PRINT_LAYOUT",
              "Apply one authoritative supported print-layout state to a sheet."
                  + " Omitted nested fields normalize to default or clear state."
                  + " The supported surface covers print area, orientation, fit scaling,"
                  + " repeating rows, repeating columns, plain header or footer text,"
                  + " margins, printGridlines, centering, paper size, draft,"
                  + " black-and-white, copies,"
                  + " first-page numbering, and explicit row or column breaks."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearPrintLayout.class,
              "CLEAR_PRINT_LAYOUT",
              "Clear the supported print-layout state from a sheet."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetCell.class, "SET_CELL", "Write one typed value to a single cell."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetRange.class,
              "SET_RANGE",
              "Write a rectangular grid of typed values."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetArrayFormula.class,
              "SET_ARRAY_FORMULA",
              "Author one contiguous single-cell or multi-cell array-formula group."
                  + " source accepts inline formula text with or without leading = or {=...}"
                  + " wrapper syntax."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearArrayFormula.class,
              "CLEAR_ARRAY_FORMULA",
              "Remove the stored array-formula group targeted by any member cell."
                  + " Non-array cells are rejected explicitly."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ImportCustomXmlMapping.class,
              "IMPORT_CUSTOM_XML_MAPPING",
              "Import XML content into one existing workbook custom-XML mapping."
                  + " mapping locates one existing map by mapId and/or name and xml accepts"
                  + " inline, file-backed, or STANDARD_INPUT text sources."
                  + " Workbook custom-XML mappings themselves must already exist in the source"
                  + " workbook; GridGrind imports data into them but does not author new map"
                  + " definitions.",
              "mapping"),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearRange.class,
              "CLEAR_RANGE",
              "Clear value, style, hyperlink, and comment state from a range."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetHyperlink.class,
              "SET_HYPERLINK",
              "Attach a hyperlink to one cell."
                  + " FILE targets use the field name path and normalize file: URIs to plain"
                  + " path strings."
                  + " Relative FILE targets are analyzed against the persisted workbook path"
                  + " when one exists."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearHyperlink.class,
              "CLEAR_HYPERLINK",
              "Remove the hyperlink from one cell; no-op when the cell does not physically exist."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetComment.class,
              "SET_COMMENT",
              "Attach a comment to one cell."
                  + " Comments can carry ordered rich-text runs and an explicit anchor box;"
                  + " runs must concatenate to text."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearComment.class,
              "CLEAR_COMMENT",
              "Remove the comment from one cell; no-op when the cell does not physically exist."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetPicture.class,
              "SET_PICTURE",
              "Create or replace one named picture on a sheet."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Reusing an existing object name replaces that picture authoritatively."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetSignatureLine.class,
              "SET_SIGNATURE_LINE",
              "Create or replace one named signature line on a sheet."
                  + " Signature lines surface through GET_DRAWING_OBJECTS and reuse"
                  + " SET_DRAWING_OBJECT_ANCHOR / DELETE_DRAWING_OBJECT for later edits."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Reusing an existing object name replaces any prior drawing object of that"
                  + " name authoritatively."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetChart.class,
              "SET_CHART",
              "Create or mutate one named chart on a sheet."
                  + " Supported authored families are AREA, AREA_3D, BAR, BAR_3D, DOUGHNUT,"
                  + " LINE, LINE_3D, PIE, PIE_3D, RADAR, SCATTER, SURFACE, and SURFACE_3D."
                  + " series bind to contiguous ranges or defined names and anchors currently"
                  + " support only TWO_CELL."
                  + " Chart and series FORMULA titles must resolve to one cell."
                  + " Failed validation leaves existing drawing state unchanged and creates no"
                  + " partial chart artifacts."
                  + " Existing unsupported chart detail is preserved on unrelated edits and"
                  + " rejected for authoritative mutation."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetPivotTable.class,
              "SET_PIVOT_TABLE",
              "Create or replace one workbook-global pivot table to POI XSSF's supported"
                  + " limited extent."
                  + " rowLabels, columnLabels, reportFilters, and dataFields must use disjoint"
                  + " source columns because POI persists only one role per pivot field."
                  + " When reportFilters are present, anchor.topLeftAddress must be on row 3 or"
                  + " lower so Excel's page-filter layout has room above the rendered body."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetShape.class,
              "SET_SHAPE",
              "Create or replace one named authored drawing shape on a sheet."
                  + " kind currently supports only SIMPLE_SHAPE and CONNECTOR."
                  + " SIMPLE_SHAPE defaults presetGeometryToken to rect when omitted."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Failed validation leaves existing drawing state unchanged and creates no"
                  + " partial shape artifacts."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetEmbeddedObject.class,
              "SET_EMBEDDED_OBJECT",
              "Create or replace one named embedded OLE package on a sheet."
                  + " previewImage supplies the visible preview raster required by Excel"
                  + " and Apache POI."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetDrawingObjectAnchor.class,
              "SET_DRAWING_OBJECT_ANCHOR",
              "Move one existing named picture, signature line, connector, simple shape,"
                  + " chart frame, or embedded object to a new authored anchor."
                  + " anchor currently supports only TWO_CELL anchors."
                  + " Read-only loaded families such as groups and graphic frames are rejected."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.DeleteDrawingObject.class,
              "DELETE_DRAWING_OBJECT",
              "Delete one existing named drawing object from the sheet."
                  + " Package relationships for picture media, signature-line preview images,"
                  + " and embedded-object parts are"
                  + " cleaned up when no other drawing object still references them."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ApplyStyle.class,
              "APPLY_STYLE",
              "Apply a style patch to every cell in a range."
                  + " Write shape: style.border is a nested object"
                  + " { \"all\": { \"style\": \"THIN\" } } or per-side top/right/bottom/left."
                  + " Read shape (GET_CELLS, GET_WINDOW): borders are flat properties"
                  + " topBorderStyle, rightBorderStyle, bottomBorderStyle, leftBorderStyle;"
                  + " the nested border object is write-only."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetDataValidation.class,
              "SET_DATA_VALIDATION",
              "Create or replace one data-validation rule over the supplied sheet range."
                  + " Overlapping existing rules are normalized so the written rule becomes"
                  + " authoritative on its target range."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearDataValidations.class,
              "CLEAR_DATA_VALIDATIONS",
              "Remove data-validation structures from the selected ranges on one sheet."
                  + " SELECTED removes only intersecting coverage; ALL clears every rule"
                  + " on the sheet."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetConditionalFormatting.class,
              "SET_CONDITIONAL_FORMATTING",
              "Create or replace one logical conditional-formatting block over the supplied"
                  + " sheet ranges."
                  + " The write contract authors formula rules, cell-value rules, color scales,"
                  + " data bars, icon sets, and top/bottom N rules."
                  + " Any existing conditional-formatting block that intersects the target ranges"
                  + " is removed first so the written block becomes authoritative on that"
                  + " coverage."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearConditionalFormatting.class,
              "CLEAR_CONDITIONAL_FORMATTING",
              "Remove conditional-formatting blocks from the selected ranges on one sheet."
                  + " SELECTED removes whole blocks whose stored ranges intersect the supplied"
                  + " ranges; ALL clears every conditional-formatting block on the sheet."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetAutofilter.class,
              "SET_AUTOFILTER",
              "Create or replace one sheet-level autofilter range."
                  + " The range must include a nonblank header row and must not overlap"
                  + " an existing table range."
                  + " criteria and sortState are optional and, when supplied, are authored"
                  + " authoritatively alongside the range.",
              "criteria",
              "sortState"),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.ClearAutofilter.class,
              "CLEAR_AUTOFILTER",
              "Clear the sheet-level autofilter range on one sheet."
                  + " Table-owned autofilters remain attached to their tables."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetTable.class,
              "SET_TABLE",
              "Create or replace one workbook-global table definition."
                  + " Table names are workbook-global and case-insensitive."
                  + " Header cells must be nonblank and unique (case-insensitive)."
                  + " Overlapping existing tables are rejected."
                  + " Any overlapping sheet-level autofilter is cleared so the table-owned"
                  + " autofilter becomes authoritative on that range."
                  + " The contract also supports advanced table metadata such as autofilter"
                  + " presence, comment, published and insert-row flags, cell-style ids,"
                  + " and per-column unique names, totals metadata, and calculated formulas."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.DeleteTable.class,
              "DELETE_TABLE",
              "Delete one existing workbook-global table by name and expected sheet name."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.DeletePivotTable.class,
              "DELETE_PIVOT_TABLE",
              "Delete one existing workbook-global pivot table by name and expected sheet name."
                  + " The expected sheet guards against accidentally deleting a same-named pivot"
                  + " after unrelated workbook changes."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.SetNamedRange.class,
              "SET_NAMED_RANGE",
              "Create or replace one workbook- or sheet-scoped named range."
                  + " target can be an explicit sheet plus A1 range, or a formula-defined target."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.DeleteNamedRange.class,
              "DELETE_NAMED_RANGE",
              "Delete one existing workbook- or sheet-scoped named range."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.AppendRow.class,
              "APPEND_ROW",
              "Append a row of typed values after the last value-bearing row."
                  + " Blank rows that carry only style, comment, or hyperlink metadata do not"
                  + " affect the append position."),
          GridGrindProtocolCatalog.descriptor(
              MutationAction.AutoSizeColumns.class,
              "AUTO_SIZE_COLUMNS",
              "Size columns deterministically from displayed cell content so the resulting"
                  + " widths are stable in headless and container runs."));

  static final List<CatalogTypeDescriptor> ASSERTION_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              Assertion.NamedRangePresent.class,
              "EXPECT_NAMED_RANGE_PRESENT",
              "Require the selected named-range selector to resolve to one or more named"
                  + " ranges."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.NamedRangeAbsent.class,
              "EXPECT_NAMED_RANGE_ABSENT",
              "Require the selected named-range selector to resolve to no named ranges."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.TablePresent.class,
              "EXPECT_TABLE_PRESENT",
              "Require the selected table selector to resolve to one or more tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.TableAbsent.class,
              "EXPECT_TABLE_ABSENT",
              "Require the selected table selector to resolve to no tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.PivotTablePresent.class,
              "EXPECT_PIVOT_TABLE_PRESENT",
              "Require the selected pivot-table selector to resolve to one or more pivot"
                  + " tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.PivotTableAbsent.class,
              "EXPECT_PIVOT_TABLE_ABSENT",
              "Require the selected pivot-table selector to resolve to no pivot tables."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.ChartPresent.class,
              "EXPECT_CHART_PRESENT",
              "Require the selected chart selector to resolve to one or more charts."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.ChartAbsent.class,
              "EXPECT_CHART_ABSENT",
              "Require the selected chart selector to resolve to no charts."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.CellValue.class,
              "EXPECT_CELL_VALUE",
              "Require every selected cell to have the exact effective value."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.DisplayValue.class,
              "EXPECT_DISPLAY_VALUE",
              "Require every selected cell to have the exact formatted display string."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.FormulaText.class,
              "EXPECT_FORMULA_TEXT",
              "Require every selected cell to store the exact formula text."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.CellStyle.class,
              "EXPECT_CELL_STYLE",
              "Require every selected cell to have the exact style snapshot."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.WorkbookProtectionFacts.class,
              "EXPECT_WORKBOOK_PROTECTION",
              "Require the workbook protection report to match exactly."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.SheetStructureFacts.class,
              "EXPECT_SHEET_STRUCTURE",
              "Require the selected sheet summary report to match exactly."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.NamedRangeFacts.class,
              "EXPECT_NAMED_RANGE_FACTS",
              "Require the selected named-range reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.TableFacts.class,
              "EXPECT_TABLE_FACTS",
              "Require the selected table reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.PivotTableFacts.class,
              "EXPECT_PIVOT_TABLE_FACTS",
              "Require the selected pivot-table reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.ChartFacts.class,
              "EXPECT_CHART_FACTS",
              "Require the selected chart reports to match exactly and in order."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnalysisMaxSeverity.class,
              "EXPECT_ANALYSIS_MAX_SEVERITY",
              "Run one analysis query against the selected target and require its highest finding"
                  + " severity to be no higher than maximumSeverity."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnalysisFindingPresent.class,
              "EXPECT_ANALYSIS_FINDING_PRESENT",
              "Run one analysis query against the selected target and require at least one"
                  + " matching finding. severity and messageContains are optional match"
                  + " refinements.",
              "severity",
              "messageContains"),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnalysisFindingAbsent.class,
              "EXPECT_ANALYSIS_FINDING_ABSENT",
              "Run one analysis query against the selected target and require no matching"
                  + " finding. severity and messageContains are optional match refinements.",
              "severity",
              "messageContains"),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AllOf.class,
              "ALL_OF",
              "Require every nested assertion to pass against the same step target."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.AnyOf.class,
              "ANY_OF",
              "Require at least one nested assertion to pass against the same step target."),
          GridGrindProtocolCatalog.descriptor(
              Assertion.Not.class,
              "NOT",
              "Invert one nested assertion against the same step target."));

  static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetWorkbookSummary.class,
              "GET_WORKBOOK_SUMMARY",
              "Return workbook-level summary facts including sheet count, sheet names,"
                  + " named range count, and formula recalculation flag."
                  + " Empty workbooks return workbook.kind=EMPTY;"
                  + " non-empty workbooks return workbook.kind=WITH_SHEETS with activeSheetName"
                  + " and selectedSheetNames."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetPackageSecurity.class,
              "GET_PACKAGE_SECURITY",
              "Return OOXML package-encryption facts and package-signature validation results."
                  + " This reports the currently loaded workbook package state;"
                  + " after in-memory mutations, previously valid source signatures read back as"
                  + " INVALIDATED_BY_MUTATION until the saved output is re-signed."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetWorkbookProtection.class,
              "GET_WORKBOOK_PROTECTION",
              "Return workbook-level protection facts including structure, windows, and"
                  + " revisions lock state plus whether password hashes are present."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetCustomXmlMappings.class,
              "GET_CUSTOM_XML_MAPPINGS",
              "Return workbook custom-XML mapping metadata, including map identifiers,"
                  + " schema metadata, linked single cells, linked tables, and optional data"
                  + " binding facts."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.ExportCustomXmlMapping.class,
              "EXPORT_CUSTOM_XML_MAPPING",
              "Export one existing workbook custom-XML mapping as serialized XML."
                  + " mapping locates one existing map by mapId and/or name;"
                  + " validateSchema defaults to false and encoding defaults to UTF-8 when"
                  + " omitted.",
              "mapping",
              "validateSchema",
              "encoding"),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetNamedRanges.class,
              "GET_NAMED_RANGES",
              "Return named ranges matched by the supplied selection."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetSheetSummary.class,
              "GET_SHEET_SUMMARY",
              "Return structural summary facts for one sheet."
                  + " Includes visibility and sheet protection state."
                  + " physicalRowCount is the number of physically materialized rows (sparse)."
                  + " lastRowIndex is the 0-based index of the last materialized row"
                  + " (-1 when empty), including metadata-only rows."
                  + " lastColumnIndex is the 0-based index of the last materialized column"
                  + " in any row (-1 when empty)."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetCells.class,
              "GET_CELLS",
              "Return exact cell snapshots for explicit addresses."
                  + " An invalid or out-of-range address returns INVALID_CELL_ADDRESS, not a blank."
                  + " Each snapshot includes address, declaredType, effectiveType, displayValue,"
                  + " style, and metadata fields. Type-specific fields: stringValue (STRING),"
                  + " numberValue (NUMBER), booleanValue (BOOLEAN), errorValue (ERROR),"
                  + " formula and evaluation (FORMULA). For FORMULA cells, effectiveType is FORMULA"
                  + " and the evaluated result type is in evaluation.effectiveType."
                  + " style.font.fontHeight is a plain object with both twips and points fields,"
                  + " nested under the style.font group."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetWindow.class,
              "GET_WINDOW",
              "Return a rectangular window of cell snapshots."
                  + " rowCount * columnCount must not exceed "
                  + InspectionQuery.MAX_WINDOW_CELLS
                  + ". Each cell snapshot has the same shape as GET_CELLS: address,"
                  + " declaredType, effectiveType, displayValue, style, metadata,"
                  + " and type-specific value fields."
                  + " For FORMULA cells, effectiveType is FORMULA and the evaluated result type"
                  + " is in evaluation.effectiveType."
                  + " Response shape: { \"window\": { \"rows\": [ { \"cells\": [...] } ] } }."
                  + " Note: the top-level key is \"window\" and cells are nested under"
                  + " window.rows[N].cells, unlike GET_CELLS which places cells directly"
                  + " under the top-level \"cells\" key."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetMergedRegions.class,
              "GET_MERGED_REGIONS",
              "Return the merged regions defined on one sheet."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetHyperlinks.class,
              "GET_HYPERLINKS",
              "Return hyperlink metadata for selected cells in the same discriminated shape used"
                  + " by SET_HYPERLINK targets. FILE targets are returned as normalized path"
                  + " strings, not file: URIs."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetComments.class,
              "GET_COMMENTS",
              "Return comment metadata for selected cells."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetDrawingObjects.class,
              "GET_DRAWING_OBJECTS",
              "Return factual drawing-object metadata for one sheet."
                  + " Read families include pictures, signature lines, simple shapes,"
                  + " connectors, groups, charts, graphic frames, and embedded objects with"
                  + " truthful anchor and package facts."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetCharts.class,
              "GET_CHARTS",
              "Return factual chart metadata for one sheet."
                  + " Supported authored chart families are modeled authoritatively;"
                  + " unsupported plot families or unsupported loaded detail are surfaced as"
                  + " explicit UNSUPPORTED entries with preserved plot-type tokens."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetArrayFormulas.class,
              "GET_ARRAY_FORMULAS",
              "Return factual array-formula group metadata for the selected sheets,"
                  + " including the stored range, top-left anchor cell, normalized formula text,"
                  + " and whether the group is single-cell."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetPivotTables.class,
              "GET_PIVOT_TABLES",
              "Return factual pivot-table metadata selected by workbook-global pivot-table name"
                  + " or ALL."
                  + " Supported pivots surface source, anchor, row or column labels,"
                  + " report filters, data fields, and values-axis placement."
                  + " Unsupported or malformed pivots are returned explicitly with preserved"
                  + " detail instead of causing read failure."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetDrawingObjectPayload.class,
              "GET_DRAWING_OBJECT_PAYLOAD",
              "Return the extracted binary payload for one existing named picture or embedded"
                  + " object."
                  + " Non-binary drawing shapes such as connectors and simple shapes are"
                  + " rejected."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetSheetLayout.class,
              "GET_SHEET_LAYOUT",
              GridGrindContractText.sheetLayoutReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetPrintLayout.class,
              "GET_PRINT_LAYOUT",
              "Return supported print-layout metadata for one sheet, including print area,"
                  + " orientation, scaling, repeating rows or columns, and plain header or"
                  + " footer text."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetDataValidations.class,
              "GET_DATA_VALIDATIONS",
              "Return factual data-validation structures for the selected sheet ranges."
                  + " Supported rules include explicit lists, formula lists, comparison rules,"
                  + " and custom formulas; unsupported rules are surfaced explicitly with typed"
                  + " detail."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetConditionalFormatting.class,
              "GET_CONDITIONAL_FORMATTING",
              "Return factual conditional-formatting blocks for the selected sheet ranges."
                  + " Read families include authored formula and cell-value rules plus loaded"
                  + " color scales, data bars, icon sets, and explicitly unsupported rules."
                  + " Each block preserves its stored ordered ranges and rule priority data."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetAutofilters.class,
              "GET_AUTOFILTERS",
              "Return sheet- and table-owned autofilter metadata for one sheet."
                  + " Each entry is typed as SHEET or TABLE so ownership is explicit."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetTables.class,
              "GET_TABLES",
              "Return factual table metadata selected by workbook-global table name or ALL."
                  + " Each table includes range, header and totals row counts, column names,"
                  + " style metadata, and whether a table-owned autofilter is present."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetFormulaSurface.class,
              "GET_FORMULA_SURFACE",
              GridGrindContractText.formulaSurfaceReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetSheetSchema.class,
              "GET_SHEET_SCHEMA",
              "Infer a simple schema from a rectangular sheet window."
                  + " rowCount * columnCount must not exceed "
                  + InspectionQuery.MAX_WINDOW_CELLS
                  + "."
                  + " The first row is treated as the header; dataRowCount is 0 when all header"
                  + " cells are blank."
                  + " dominantType is null when all data cells are blank, or when two or more types"
                  + " tie for the highest count."
                  + " Formula cells contribute their evaluated result type (NUMBER, STRING, etc.)"
                  + " to dominantType and observedTypes, not FORMULA."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.GetNamedRangeSurface.class,
              "GET_NAMED_RANGE_SURFACE",
              GridGrindContractText.namedRangeSurfaceReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeFormulaHealth.class,
              "ANALYZE_FORMULA_HEALTH",
              GridGrindContractText.formulaHealthReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeDataValidationHealth.class,
              "ANALYZE_DATA_VALIDATION_HEALTH",
              "Return analysis.checkedValidationCount, a severity summary,"
                  + " and findings such as unsupported, overlapping, or"
                  + " broken-formula rules."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeConditionalFormattingHealth.class,
              "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
              "Return analysis.checkedConditionalFormattingBlockCount,"
                  + " a severity summary, and conditional-formatting findings such as broken formulas,"
                  + " unsupported loaded rules, invalid target ranges, or priority collisions."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeAutofilterHealth.class,
              "ANALYZE_AUTOFILTER_HEALTH",
              "Return analysis.checkedAutofilterCount, a severity summary,"
                  + " and autofilter findings such as invalid ranges, blank header rows,"
                  + " or ownership mismatches between sheet-level filters and tables."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeTableHealth.class,
              "ANALYZE_TABLE_HEALTH",
              "Return analysis.checkedTableCount, a severity summary,"
                  + " and table findings such as overlaps, broken references,"
                  + " blank or duplicate headers, and unresolved styles."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzePivotTableHealth.class,
              "ANALYZE_PIVOT_TABLE_HEALTH",
              "Return analysis.checkedPivotTableCount, a severity summary,"
                  + " and pivot-table findings such as missing cache parts, broken sources,"
                  + " duplicate or synthetic names, and unsupported persisted detail."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeHyperlinkHealth.class,
              "ANALYZE_HYPERLINK_HEALTH",
              "Return analysis.checkedHyperlinkCount, a severity summary,"
                  + " and hyperlink findings such as malformed, missing, unresolved, or broken"
                  + " targets."
                  + " Relative FILE targets are resolved against the workbook's persisted path"
                  + " when one exists."),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeNamedRangeHealth.class,
              "ANALYZE_NAMED_RANGE_HEALTH",
              GridGrindContractText.namedRangeHealthReadSummary()),
          GridGrindProtocolCatalog.descriptor(
              InspectionQuery.AnalyzeWorkbookFindings.class,
              "ANALYZE_WORKBOOK_FINDINGS",
              GridGrindContractText.workbookFindingsReadSummary()));

  private GridGrindProtocolCatalogTypeDescriptors() {}
}
