package dev.erst.gridgrind.excel;

/** Resolves canonical operation-style discriminator strings for workbook commands. */
final class WorkbookCommandTypeResolver {
  private WorkbookCommandTypeResolver() {}

  static String commandType(WorkbookCommand command) {
    return switch (command) {
      case WorkbookSheetCommand.CreateSheet _ -> "ENSURE_SHEET";
      case WorkbookSheetCommand.RenameSheet _ -> "RENAME_SHEET";
      case WorkbookSheetCommand.DeleteSheet _ -> "DELETE_SHEET";
      case WorkbookSheetCommand.MoveSheet _ -> "MOVE_SHEET";
      case WorkbookSheetCommand.CopySheet _ -> "COPY_SHEET";
      case WorkbookSheetCommand.SetActiveSheet _ -> "SET_ACTIVE_SHEET";
      case WorkbookSheetCommand.SetSelectedSheets _ -> "SET_SELECTED_SHEETS";
      case WorkbookSheetCommand.SetSheetVisibility _ -> "SET_SHEET_VISIBILITY";
      case WorkbookSheetCommand.SetSheetProtection _ -> "SET_SHEET_PROTECTION";
      case WorkbookSheetCommand.ClearSheetProtection _ -> "CLEAR_SHEET_PROTECTION";
      case WorkbookSheetCommand.SetWorkbookProtection _ -> "SET_WORKBOOK_PROTECTION";
      case WorkbookSheetCommand.ClearWorkbookProtection _ -> "CLEAR_WORKBOOK_PROTECTION";
      case WorkbookStructureCommand.MergeCells _ -> "MERGE_CELLS";
      case WorkbookStructureCommand.UnmergeCells _ -> "UNMERGE_CELLS";
      case WorkbookStructureCommand.SetColumnWidth _ -> "SET_COLUMN_WIDTH";
      case WorkbookStructureCommand.SetRowHeight _ -> "SET_ROW_HEIGHT";
      case WorkbookStructureCommand.InsertRows _ -> "INSERT_ROWS";
      case WorkbookStructureCommand.DeleteRows _ -> "DELETE_ROWS";
      case WorkbookStructureCommand.ShiftRows _ -> "SHIFT_ROWS";
      case WorkbookStructureCommand.InsertColumns _ -> "INSERT_COLUMNS";
      case WorkbookStructureCommand.DeleteColumns _ -> "DELETE_COLUMNS";
      case WorkbookStructureCommand.ShiftColumns _ -> "SHIFT_COLUMNS";
      case WorkbookStructureCommand.SetRowVisibility _ -> "SET_ROW_VISIBILITY";
      case WorkbookStructureCommand.SetColumnVisibility _ -> "SET_COLUMN_VISIBILITY";
      case WorkbookStructureCommand.GroupRows _ -> "GROUP_ROWS";
      case WorkbookStructureCommand.UngroupRows _ -> "UNGROUP_ROWS";
      case WorkbookStructureCommand.GroupColumns _ -> "GROUP_COLUMNS";
      case WorkbookStructureCommand.UngroupColumns _ -> "UNGROUP_COLUMNS";
      case WorkbookLayoutCommand.SetSheetPane _ -> "SET_SHEET_PANE";
      case WorkbookLayoutCommand.SetSheetZoom _ -> "SET_SHEET_ZOOM";
      case WorkbookLayoutCommand.SetSheetPresentation _ -> "SET_SHEET_PRESENTATION";
      case WorkbookLayoutCommand.SetPrintLayout _ -> "SET_PRINT_LAYOUT";
      case WorkbookLayoutCommand.ClearPrintLayout _ -> "CLEAR_PRINT_LAYOUT";
      case WorkbookCellCommand.SetCell _ -> "SET_CELL";
      case WorkbookCellCommand.SetRange _ -> "SET_RANGE";
      case WorkbookCellCommand.ClearRange _ -> "CLEAR_RANGE";
      case WorkbookCellCommand.SetArrayFormula _ -> "SET_ARRAY_FORMULA";
      case WorkbookCellCommand.ClearArrayFormula _ -> "CLEAR_ARRAY_FORMULA";
      case WorkbookMetadataCommand.ImportCustomXmlMapping _ -> "IMPORT_CUSTOM_XML_MAPPING";
      case WorkbookAnnotationCommand.SetHyperlink _ -> "SET_HYPERLINK";
      case WorkbookAnnotationCommand.ClearHyperlink _ -> "CLEAR_HYPERLINK";
      case WorkbookAnnotationCommand.SetComment _ -> "SET_COMMENT";
      case WorkbookAnnotationCommand.ClearComment _ -> "CLEAR_COMMENT";
      case WorkbookDrawingCommand.SetPicture _ -> "SET_PICTURE";
      case WorkbookDrawingCommand.SetSignatureLine _ -> "SET_SIGNATURE_LINE";
      case WorkbookDrawingCommand.SetChart _ -> "SET_CHART";
      case WorkbookTabularCommand.SetPivotTable _ -> "SET_PIVOT_TABLE";
      case WorkbookDrawingCommand.SetShape _ -> "SET_SHAPE";
      case WorkbookDrawingCommand.SetEmbeddedObject _ -> "SET_EMBEDDED_OBJECT";
      case WorkbookDrawingCommand.SetDrawingObjectAnchor _ -> "SET_DRAWING_OBJECT_ANCHOR";
      case WorkbookDrawingCommand.DeleteDrawingObject _ -> "DELETE_DRAWING_OBJECT";
      case WorkbookFormattingCommand.ApplyStyle _ -> "APPLY_STYLE";
      case WorkbookFormattingCommand.SetDataValidation _ -> "SET_DATA_VALIDATION";
      case WorkbookFormattingCommand.ClearDataValidations _ -> "CLEAR_DATA_VALIDATIONS";
      case WorkbookFormattingCommand.SetConditionalFormatting _ -> "SET_CONDITIONAL_FORMATTING";
      case WorkbookFormattingCommand.ClearConditionalFormatting _ -> "CLEAR_CONDITIONAL_FORMATTING";
      case WorkbookTabularCommand.SetAutofilter _ -> "SET_AUTOFILTER";
      case WorkbookTabularCommand.ClearAutofilter _ -> "CLEAR_AUTOFILTER";
      case WorkbookTabularCommand.SetTable _ -> "SET_TABLE";
      case WorkbookTabularCommand.DeleteTable _ -> "DELETE_TABLE";
      case WorkbookTabularCommand.DeletePivotTable _ -> "DELETE_PIVOT_TABLE";
      case WorkbookMetadataCommand.SetNamedRange _ -> "SET_NAMED_RANGE";
      case WorkbookMetadataCommand.DeleteNamedRange _ -> "DELETE_NAMED_RANGE";
      case WorkbookCellCommand.AppendRow _ -> "APPEND_ROW";
      case WorkbookLayoutCommand.AutoSizeColumns _ -> "AUTO_SIZE_COLUMNS";
    };
  }
}
