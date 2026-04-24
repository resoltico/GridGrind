package dev.erst.gridgrind.excel;

/** Resolves canonical operation-style discriminator strings for workbook commands. */
final class WorkbookCommandTypeResolver {
  private WorkbookCommandTypeResolver() {}

  static String commandType(WorkbookCommand command) {
    return switch (command) {
      case WorkbookCommand.CreateSheet _ -> "ENSURE_SHEET";
      case WorkbookCommand.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookCommand.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookCommand.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookCommand.CopySheet _ -> "COPY_SHEET";
      case WorkbookCommand.SetActiveSheet _ -> "SET_ACTIVE_SHEET";
      case WorkbookCommand.SetSelectedSheets _ -> "SET_SELECTED_SHEETS";
      case WorkbookCommand.SetSheetVisibility _ -> "SET_SHEET_VISIBILITY";
      case WorkbookCommand.SetSheetProtection _ -> "SET_SHEET_PROTECTION";
      case WorkbookCommand.ClearSheetProtection _ -> "CLEAR_SHEET_PROTECTION";
      case WorkbookCommand.SetWorkbookProtection _ -> "SET_WORKBOOK_PROTECTION";
      case WorkbookCommand.ClearWorkbookProtection _ -> "CLEAR_WORKBOOK_PROTECTION";
      case WorkbookCommand.MergeCells _ -> "MERGE_CELLS";
      case WorkbookCommand.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookCommand.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookCommand.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookCommand.InsertRows _ -> "INSERT_ROWS";
      case WorkbookCommand.DeleteRows _ -> "DELETE_ROWS";
      case WorkbookCommand.ShiftRows _ -> "SHIFT_ROWS";
      case WorkbookCommand.InsertColumns _ -> "INSERT_COLUMNS";
      case WorkbookCommand.DeleteColumns _ -> "DELETE_COLUMNS";
      case WorkbookCommand.ShiftColumns _ -> "SHIFT_COLUMNS";
      case WorkbookCommand.SetRowVisibility _ -> "SET_ROW_VISIBILITY";
      case WorkbookCommand.SetColumnVisibility _ -> "SET_COLUMN_VISIBILITY";
      case WorkbookCommand.GroupRows _ -> "GROUP_ROWS";
      case WorkbookCommand.UngroupRows _ -> "UNGROUP_ROWS";
      case WorkbookCommand.GroupColumns _ -> "GROUP_COLUMNS";
      case WorkbookCommand.UngroupColumns _ -> "UNGROUP_COLUMNS";
      case WorkbookCommand.SetSheetPane _ -> "SET_SHEET_PANE";
      case WorkbookCommand.SetSheetZoom _ -> "SET_SHEET_ZOOM";
      case WorkbookCommand.SetSheetPresentation _ -> "SET_SHEET_PRESENTATION";
      case WorkbookCommand.SetPrintLayout _ -> "SET_PRINT_LAYOUT";
      case WorkbookCommand.ClearPrintLayout _ -> "CLEAR_PRINT_LAYOUT";
      case WorkbookCommand.SetCell _ -> "SET_CELL";
      case WorkbookCommand.SetRange _ -> "SET_RANGE";
      case WorkbookCommand.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookCommand.SetArrayFormula _ -> "SET_ARRAY_FORMULA";
      case WorkbookCommand.ClearArrayFormula _ -> "CLEAR_ARRAY_FORMULA";
      case WorkbookCommand.ImportCustomXmlMapping _ -> "IMPORT_CUSTOM_XML_MAPPING";
      case WorkbookCommand.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookCommand.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookCommand.SetComment _ -> "SET_COMMENT";
      case WorkbookCommand.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookCommand.SetPicture _ -> "SET_PICTURE";
      case WorkbookCommand.SetSignatureLine _ -> "SET_SIGNATURE_LINE";
      case WorkbookCommand.SetChart _ -> "SET_CHART";
      case WorkbookCommand.SetPivotTable _ -> "SET_PIVOT_TABLE";
      case WorkbookCommand.SetShape _ -> "SET_SHAPE";
      case WorkbookCommand.SetEmbeddedObject _ -> "SET_EMBEDDED_OBJECT";
      case WorkbookCommand.SetDrawingObjectAnchor _ -> "SET_DRAWING_OBJECT_ANCHOR";
      case WorkbookCommand.DeleteDrawingObject _ -> "DELETE_DRAWING_OBJECT";
      case WorkbookCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookCommand.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookCommand.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookCommand.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookCommand.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookCommand.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookCommand.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookCommand.SetTable _ -> "SET_TABLE";
      case WorkbookCommand.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookCommand.DeletePivotTable _ -> "DELETE_PIVOT_TABLE";
      case WorkbookCommand.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookCommand.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
    };
  }
}
