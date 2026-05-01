package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import java.util.List;

/** Mutation-action descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogMutationActionTypeDescriptors {
  static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.EnsureSheet.class,
              "ENSURE_SHEET",
              "Create the sheet if it does not already exist."
                  + " Sheet names must be 1 to 31 characters and must not contain"
                  + " : \\ / ? * [ ] or begin or end with a single quote (Excel limit)."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.RenameSheet.class,
              "RENAME_SHEET",
              "Rename an existing sheet."
                  + " The new name must be 1 to 31 characters and must not contain"
                  + " : \\ / ? * [ ] or begin or end with a single quote (Excel limit)."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.DeleteSheet.class,
              "DELETE_SHEET",
              "Delete an existing sheet."
                  + " A workbook must retain at least one sheet and at least one visible sheet;"
                  + " deleting the last sheet or the last visible sheet returns INVALID_REQUEST."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.MoveSheet.class,
              "MOVE_SHEET",
              "Move a sheet to a zero-based workbook position."
                  + " targetIndex is 0-based: 0 moves the sheet to the front,"
                  + " sheetCount-1 moves it to the back."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.CopySheet.class,
              "COPY_SHEET",
              "Copy one sheet into a new visible, unselected sheet."
                  + " newSheetName follows the same Excel sheet-name rules as ENSURE_SHEET."
                  + " position is explicit on the wire; the Java authoring surface exposes"
                  + " one append-at-end convenience constructor."
                  + " The copied sheet preserves sheet-local workbook content such as formulas,"
                  + " validations, conditional formatting, comments, hyperlinks, merged regions,"
                  + " layout state, sheet-scoped named ranges, sheet autofilters, and tables."
                  + " Copied tables are renamed automatically when needed to keep workbook-global"
                  + " table names unique.",
              "position"),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetActiveSheet.class,
              "SET_ACTIVE_SHEET",
              "Set the active sheet."
                  + " Hidden sheets cannot be activated."
                  + " The active sheet is always selected."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetSelectedSheets.class,
              "SET_SELECTED_SHEETS",
              "Set the selected visible sheet set."
                  + " Duplicate or unknown sheet names are rejected."
                  + " selectedSheetNames read back in workbook order;"
                  + " activeSheetName preserves the primary selected sheet choice."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetSheetVisibility.class,
              "SET_SHEET_VISIBILITY",
              "Set one sheet visibility state."
                  + " A workbook must retain at least one visible sheet;"
                  + " hiding the last visible sheet is rejected."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetSheetProtection.class,
              "SET_SHEET_PROTECTION",
              "Enable sheet protection with the exact supported lock flags."
                  + " password is optional; when provided it is hashed into the sheet-protection"
                  + " metadata.",
              "password"),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.ClearSheetProtection.class,
              "CLEAR_SHEET_PROTECTION",
              "Disable sheet protection entirely."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetWorkbookProtection.class,
              "SET_WORKBOOK_PROTECTION",
              "Enable workbook-level protection and optional workbook or revisions passwords."
                  + " Omitted lock flags normalize to false and omitted passwords are cleared."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.ClearWorkbookProtection.class,
              "CLEAR_WORKBOOK_PROTECTION",
              "Clear workbook-level protection and any stored workbook or revisions passwords."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.MergeCells.class,
              "MERGE_CELLS",
              "Merge a rectangular A1-style range."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.UnmergeCells.class,
              "UNMERGE_CELLS",
              "Remove one merged region by exact range match."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetColumnWidth.class,
              "SET_COLUMN_WIDTH",
              "Set one or more column widths in Excel character units."
                  + " widthCharacters must be > 0 and <= 255 (Excel column width limit)."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetRowHeight.class,
              "SET_ROW_HEIGHT",
              "Set one or more row heights in Excel point units."
                  + " heightPoints must be > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + " (Excel row height limit)."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.InsertRows.class,
              "INSERT_ROWS",
              "Insert one or more blank rows before rowIndex."
                  + " rowIndex must be <= last existing row + 1."
                  + " Append-edge inserts on sparse sheets do not materialize a new physical tail"
                  + " row until content or row metadata exists there."
                  + " GridGrind preserves and retargets existing data validations across the"
                  + " insert, but still rejects row inserts that would move tables or"
                  + " sheet autofilters."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.DeleteRows.class,
              "DELETE_ROWS",
              "Delete one inclusive zero-based row band."
                  + " GridGrind rejects row deletes that would move or truncate tables,"
                  + " sheet autofilters, or data validations, and it also rejects deletes that"
                  + " would truncate range-backed named ranges."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.ShiftRows.class,
              "SHIFT_ROWS",
              "Move one inclusive zero-based row band by delta rows."
                  + " delta must not be 0."
                  + " GridGrind rejects row shifts that would move tables, sheet autofilters,"
                  + " or data validations, and it also rejects shifts that would partially move"
                  + " or overwrite range-backed named ranges."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.InsertColumns.class,
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
              WorkbookMutationAction.DeleteColumns.class,
              "DELETE_COLUMNS",
              "Delete one inclusive zero-based column band."
                  + " GridGrind rejects column deletes when the workbook contains formula cells"
                  + " or formula-defined names, or when the edit would move or truncate tables,"
                  + " sheet autofilters, or data validations, or truncate range-backed named"
                  + " ranges."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.ShiftColumns.class,
              "SHIFT_COLUMNS",
              "Move one inclusive zero-based column band by delta columns."
                  + " delta must not be 0."
                  + " GridGrind rejects column shifts when the workbook contains formula cells"
                  + " or formula-defined names, or when the edit would move tables,"
                  + " sheet autofilters, or data validations, or partially move or overwrite"
                  + " range-backed named ranges."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetRowVisibility.class,
              "SET_ROW_VISIBILITY",
              "Set the hidden state for one inclusive zero-based row band."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetColumnVisibility.class,
              "SET_COLUMN_VISIBILITY",
              "Set the hidden state for one inclusive zero-based column band."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.GroupRows.class,
              "GROUP_ROWS",
              "Apply one outline group to an inclusive zero-based row band."
                  + " collapsed is explicit on the wire; use GroupRows.expanded() from Java"
                  + " authoring when the expanded form is intended."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.UngroupRows.class,
              "UNGROUP_ROWS",
              "Remove outline grouping from one inclusive zero-based row band."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.GroupColumns.class,
              "GROUP_COLUMNS",
              "Apply one outline group to an inclusive zero-based column band."
                  + " collapsed is explicit on the wire; use GroupColumns.expanded() from Java"
                  + " authoring when the expanded form is intended."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.UngroupColumns.class,
              "UNGROUP_COLUMNS",
              "Remove outline grouping from one inclusive zero-based column band."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetSheetPane.class,
              "SET_SHEET_PANE",
              "Apply one explicit pane state."
                  + " pane.type can be NONE, FROZEN, or SPLIT; use NONE to clear panes."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetSheetZoom.class,
              "SET_SHEET_ZOOM",
              "Set the sheet zoom percentage."
                  + " zoomPercent must be between "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MIN_ZOOM_PERCENT
                  + " and "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ZOOM_PERCENT
                  + " inclusive."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetSheetPresentation.class,
              "SET_SHEET_PRESENTATION",
              "Apply one authoritative supported sheet-presentation state to a sheet."
                  + " Supply explicit display, tabColor, outlineSummary, sheetDefaults, and"
                  + " ignoredErrors values, or use SheetPresentationInput.defaults() from Java"
                  + " authoring."
                  + " The supported surface covers screen display flags, right-to-left mode,"
                  + " tab color, outline summary placement, authored default row and column sizing"
                  + " (defaultColumnWidth > 0 and <= 255; defaultRowHeightPoints > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + "), and ignored-errors ranges."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.SetPrintLayout.class,
              "SET_PRINT_LAYOUT",
              "Apply one authoritative supported print-layout state to a sheet."
                  + " Supply explicit printArea, orientation, scaling, repeating rows,"
                  + " repeating columns, header, footer, and setup values, or use"
                  + " PrintLayoutInput.defaults() / withDefaultSetup(...) from Java authoring."
                  + " The supported surface covers print area, orientation, fit scaling,"
                  + " repeating rows, repeating columns, plain header or footer text,"
                  + " margins, printGridlines, centering, paper size, draft,"
                  + " black-and-white, copies,"
                  + " first-page numbering, and explicit row or column breaks."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.ClearPrintLayout.class,
              "CLEAR_PRINT_LAYOUT",
              "Clear the supported print-layout state from a sheet."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.SetCell.class,
              "SET_CELL",
              "Write one typed value to a single cell."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.SetRange.class,
              "SET_RANGE",
              "Write a rectangular grid of typed values."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.SetArrayFormula.class,
              "SET_ARRAY_FORMULA",
              "Author one contiguous single-cell or multi-cell array-formula group."
                  + " source accepts inline formula text with or without leading = or {=...}"
                  + " wrapper syntax."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.ClearArrayFormula.class,
              "CLEAR_ARRAY_FORMULA",
              "Remove the stored array-formula group targeted by any member cell."
                  + " Non-array cells are rejected explicitly."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.ImportCustomXmlMapping.class,
              "IMPORT_CUSTOM_XML_MAPPING",
              "Import XML content into one existing workbook custom-XML mapping."
                  + " mapping locates one existing map by mapId and/or name and xml accepts"
                  + " inline, file-backed, or STANDARD_INPUT text sources."
                  + " Workbook custom-XML mappings themselves must already exist in the source"
                  + " workbook; GridGrind imports data into them but does not author new map"
                  + " definitions.",
              "mapping"),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.ClearRange.class,
              "CLEAR_RANGE",
              "Clear value, style, hyperlink, and comment state from a range."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.SetHyperlink.class,
              "SET_HYPERLINK",
              "Attach a hyperlink to one cell."
                  + " FILE targets use the field name path and normalize file: URIs to plain"
                  + " path strings."
                  + " Relative FILE targets are analyzed against the persisted workbook path"
                  + " when one exists."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.ClearHyperlink.class,
              "CLEAR_HYPERLINK",
              "Remove the hyperlink from one cell; no-op when the cell does not physically exist."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.SetComment.class,
              "SET_COMMENT",
              "Attach a comment to one cell."
                  + " Comments can carry ordered rich-text runs and an explicit anchor box;"
                  + " runs must concatenate to text."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.ClearComment.class,
              "CLEAR_COMMENT",
              "Remove the comment from one cell; no-op when the cell does not physically exist."),
          GridGrindProtocolCatalog.descriptor(
              DrawingMutationAction.SetPicture.class,
              "SET_PICTURE",
              "Create or replace one named picture on a sheet."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Reusing an existing object name replaces that picture authoritatively."),
          GridGrindProtocolCatalog.descriptor(
              DrawingMutationAction.SetSignatureLine.class,
              "SET_SIGNATURE_LINE",
              "Create or replace one named signature line on a sheet."
                  + " Signature lines surface through GET_DRAWING_OBJECTS and reuse"
                  + " SET_DRAWING_OBJECT_ANCHOR / DELETE_DRAWING_OBJECT for later edits."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Reusing an existing object name replaces any prior drawing object of that"
                  + " name authoritatively."),
          GridGrindProtocolCatalog.descriptor(
              DrawingMutationAction.SetChart.class,
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
              StructuredMutationAction.SetPivotTable.class,
              "SET_PIVOT_TABLE",
              "Create or replace one workbook-global pivot table to POI XSSF's supported"
                  + " limited extent."
                  + " rowLabels, columnLabels, reportFilters, and dataFields must use disjoint"
                  + " source columns because POI persists only one role per pivot field."
                  + " When reportFilters are present, anchor.topLeftAddress must be on row 3 or"
                  + " lower so Excel's page-filter layout has room above the rendered body."),
          GridGrindProtocolCatalog.descriptor(
              DrawingMutationAction.SetShape.class,
              "SET_SHAPE",
              "Create or replace one named authored drawing shape on a sheet."
                  + " kind currently supports only SIMPLE_SHAPE and CONNECTOR."
                  + " SIMPLE_SHAPE defaults presetGeometryToken to rect when omitted."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Failed validation leaves existing drawing state unchanged and creates no"
                  + " partial shape artifacts."),
          GridGrindProtocolCatalog.descriptor(
              DrawingMutationAction.SetEmbeddedObject.class,
              "SET_EMBEDDED_OBJECT",
              "Create or replace one named embedded OLE package on a sheet."
                  + " previewImage supplies the visible preview raster required by Excel"
                  + " and Apache POI."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."),
          GridGrindProtocolCatalog.descriptor(
              DrawingMutationAction.SetDrawingObjectAnchor.class,
              "SET_DRAWING_OBJECT_ANCHOR",
              "Move one existing named picture, signature line, connector, simple shape,"
                  + " chart frame, or embedded object to a new authored anchor."
                  + " anchor currently supports only TWO_CELL anchors."
                  + " Read-only loaded families such as groups and graphic frames are rejected."),
          GridGrindProtocolCatalog.descriptor(
              DrawingMutationAction.DeleteDrawingObject.class,
              "DELETE_DRAWING_OBJECT",
              "Delete one existing named drawing object from the sheet."
                  + " Package relationships for picture media, signature-line preview images,"
                  + " and embedded-object parts are"
                  + " cleaned up when no other drawing object still references them."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.ApplyStyle.class,
              "APPLY_STYLE",
              "Apply a style patch to every cell in a range."
                  + " Write shape: style.border is a nested object"
                  + " { \"all\": { \"style\": \"THIN\" } } or per-side top/right/bottom/left."
                  + " Read shape (GET_CELLS, GET_WINDOW): borders are flat properties"
                  + " topBorderStyle, rightBorderStyle, bottomBorderStyle, leftBorderStyle;"
                  + " the nested border object is write-only."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.SetDataValidation.class,
              "SET_DATA_VALIDATION",
              "Create or replace one data-validation rule over the supplied sheet range."
                  + " Overlapping existing rules are normalized so the written rule becomes"
                  + " authoritative on its target range."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.ClearDataValidations.class,
              "CLEAR_DATA_VALIDATIONS",
              "Remove data-validation structures from the selected ranges on one sheet."
                  + " SELECTED removes only intersecting coverage; ALL clears every rule"
                  + " on the sheet."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.SetConditionalFormatting.class,
              "SET_CONDITIONAL_FORMATTING",
              "Create or replace one logical conditional-formatting block over the supplied"
                  + " sheet ranges."
                  + " The write contract authors formula rules, cell-value rules, color scales,"
                  + " data bars, icon sets, and top/bottom N rules."
                  + " Any existing conditional-formatting block that intersects the target ranges"
                  + " is removed first so the written block becomes authoritative on that"
                  + " coverage."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.ClearConditionalFormatting.class,
              "CLEAR_CONDITIONAL_FORMATTING",
              "Remove conditional-formatting blocks from the selected ranges on one sheet."
                  + " SELECTED removes whole blocks whose stored ranges intersect the supplied"
                  + " ranges; ALL clears every conditional-formatting block on the sheet."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.SetAutofilter.class,
              "SET_AUTOFILTER",
              "Create or replace one sheet-level autofilter range."
                  + " The range must include a nonblank header row and must not overlap"
                  + " an existing table range."
                  + " criteria and sortState are optional and, when supplied, are authored"
                  + " authoritatively alongside the range.",
              "criteria",
              "sortState"),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.ClearAutofilter.class,
              "CLEAR_AUTOFILTER",
              "Clear the sheet-level autofilter range on one sheet."
                  + " Table-owned autofilters remain attached to their tables."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.SetTable.class,
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
              StructuredMutationAction.DeleteTable.class,
              "DELETE_TABLE",
              "Delete one existing workbook-global table by name and expected sheet name."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.DeletePivotTable.class,
              "DELETE_PIVOT_TABLE",
              "Delete one existing workbook-global pivot table by name and expected sheet name."
                  + " The expected sheet guards against accidentally deleting a same-named pivot"
                  + " after unrelated workbook changes."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.SetNamedRange.class,
              "SET_NAMED_RANGE",
              "Create or replace one workbook- or sheet-scoped named range."
                  + " target can be an explicit sheet plus A1 range, or a formula-defined target."),
          GridGrindProtocolCatalog.descriptor(
              StructuredMutationAction.DeleteNamedRange.class,
              "DELETE_NAMED_RANGE",
              "Delete one existing workbook- or sheet-scoped named range."),
          GridGrindProtocolCatalog.descriptor(
              CellMutationAction.AppendRow.class,
              "APPEND_ROW",
              "Append a row of typed values after the last value-bearing row."
                  + " Blank rows that carry only style, comment, or hyperlink metadata do not"
                  + " affect the append position."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookMutationAction.AutoSizeColumns.class,
              "AUTO_SIZE_COLUMNS",
              "Size columns deterministically from displayed cell content so the resulting"
                  + " widths are stable in headless and container runs."));

  private GridGrindProtocolCatalogMutationActionTypeDescriptors() {}
}
