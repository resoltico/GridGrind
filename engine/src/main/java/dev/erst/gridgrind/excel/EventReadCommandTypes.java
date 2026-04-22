package dev.erst.gridgrind.excel;

/** Stable EVENT_READ command-type labelling for diagnostics and test coverage. */
final class EventReadCommandTypes {
  private EventReadCommandTypes() {}

  static String commandType(WorkbookReadCommand.Introspection command) {
    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadCommand.GetWorkbookProtection _ -> "GET_WORKBOOK_PROTECTION";
      case WorkbookReadCommand.GetCustomXmlMappings _ -> "GET_CUSTOM_XML_MAPPINGS";
      case WorkbookReadCommand.ExportCustomXmlMapping _ -> "EXPORT_CUSTOM_XML_MAPPING";
      case WorkbookReadCommand.GetNamedRanges _ -> "GET_NAMED_RANGES";
      case WorkbookReadCommand.GetSheetSummary _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadCommand.GetArrayFormulas _ -> "GET_ARRAY_FORMULAS";
      case WorkbookReadCommand.GetCells _ -> "GET_CELLS";
      case WorkbookReadCommand.GetWindow _ -> "GET_WINDOW";
      case WorkbookReadCommand.GetMergedRegions _ -> "GET_MERGED_REGIONS";
      case WorkbookReadCommand.GetHyperlinks _ -> "GET_HYPERLINKS";
      case WorkbookReadCommand.GetComments _ -> "GET_COMMENTS";
      case WorkbookReadCommand.GetDrawingObjects _ -> "GET_DRAWING_OBJECTS";
      case WorkbookReadCommand.GetCharts _ -> "GET_CHARTS";
      case WorkbookReadCommand.GetPivotTables _ -> "GET_PIVOT_TABLES";
      case WorkbookReadCommand.GetDrawingObjectPayload _ -> "GET_DRAWING_OBJECT_PAYLOAD";
      case WorkbookReadCommand.GetSheetLayout _ -> "GET_SHEET_LAYOUT";
      case WorkbookReadCommand.GetPrintLayout _ -> "GET_PRINT_LAYOUT";
      case WorkbookReadCommand.GetDataValidations _ -> "GET_DATA_VALIDATIONS";
      case WorkbookReadCommand.GetConditionalFormatting _ -> "GET_CONDITIONAL_FORMATTING";
      case WorkbookReadCommand.GetAutofilters _ -> "GET_AUTOFILTERS";
      case WorkbookReadCommand.GetTables _ -> "GET_TABLES";
      case WorkbookReadCommand.GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadCommand.GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadCommand.GetPackageSecurity _ -> "GET_PACKAGE_SECURITY";
      case WorkbookReadCommand.GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
    };
  }
}
