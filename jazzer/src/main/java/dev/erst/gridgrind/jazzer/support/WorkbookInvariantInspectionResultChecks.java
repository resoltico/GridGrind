package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks.require;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;

/** Owns inspection-query to inspection-result matching invariants for protocol workflows. */
final class WorkbookInvariantInspectionResultChecks {
  private WorkbookInvariantInspectionResultChecks() {}

  static void requireReadMatchesRequest(InspectionStep readOperation, InspectionResult readResult) {
    require(
        readOperation.stepId().equals(readResult.stepId()),
        "read result stepId must match the request");
    require(
        SequenceIntrospection.inspectionKind(readOperation).equals(readResultKind(readResult)),
        "read result kind must match the requested read kind");

    switch (readOperation.query()) {
      case InspectionQuery.GetWorkbookSummary _ -> {
        InspectionResult.WorkbookSummaryResult result =
            (InspectionResult.WorkbookSummaryResult) readResult;
        WorkbookInvariantResponseChecks.requireWorkbookSummaryShape(result.workbook());
      }
      case InspectionQuery.GetPackageSecurity _ ->
          WorkbookInvariantResponseChecks.requirePackageSecurityShape(
              ((InspectionResult.PackageSecurityResult) readResult).security());
      case InspectionQuery.GetWorkbookProtection _ ->
          WorkbookInvariantResponseChecks.requireWorkbookProtectionShape(
              ((InspectionResult.WorkbookProtectionResult) readResult).protection());
      case InspectionQuery.GetCustomXmlMappings _ -> {
        InspectionResult.CustomXmlMappingsResult result =
            (InspectionResult.CustomXmlMappingsResult) readResult;
        result.mappings().forEach(WorkbookInvariantResponseChecks::requireCustomXmlMappingShape);
      }
      case InspectionQuery.ExportCustomXmlMapping _ ->
          WorkbookInvariantResponseChecks.requireCustomXmlExportShape(
              ((InspectionResult.CustomXmlExportResult) readResult).export());
      case InspectionQuery.GetNamedRanges _ -> {
        InspectionResult.NamedRangesResult result = (InspectionResult.NamedRangesResult) readResult;
        result.namedRanges().forEach(WorkbookInvariantResponseChecks::requireNamedRangeShape);
      }
      case InspectionQuery.GetSheetSummary _ -> {
        InspectionResult.SheetSummaryResult result =
            (InspectionResult.SheetSummaryResult) readResult;
        WorkbookInvariantResponseChecks.requireSheetSummaryShape(result.sheet());
        require(
            ((SheetSelector.ByName) readOperation.target())
                .name()
                .equals(result.sheet().sheetName()),
            "sheet summary sheet mismatch");
      }
      case InspectionQuery.GetArrayFormulas _ -> {
        InspectionResult.ArrayFormulasResult result =
            (InspectionResult.ArrayFormulasResult) readResult;
        result.arrayFormulas().forEach(WorkbookInvariantResponseChecks::requireArrayFormulaShape);
      }
      case InspectionQuery.GetCells _ -> {
        InspectionResult.CellsResult result = (InspectionResult.CellsResult) readResult;
        require(
            sheetName((CellSelector) readOperation.target()).equals(result.sheetName()),
            "cells sheet mismatch");
        if (readOperation.target() instanceof CellSelector.ByAddresses byAddresses) {
          require(
              result.cells().size() == byAddresses.addresses().size(),
              "cells result size must match requested addresses");
        } else if (readOperation.target() instanceof CellSelector.ByAddress) {
          require(result.cells().size() == 1, "single-cell result size must be 1");
        }
      }
      case InspectionQuery.GetWindow _ -> {
        InspectionResult.WindowResult result = (InspectionResult.WindowResult) readResult;
        RangeSelector.RectangularWindow selector =
            (RangeSelector.RectangularWindow) readOperation.target();
        require(selector.sheetName().equals(result.window().sheetName()), "window sheet mismatch");
        require(
            selector.topLeftAddress().equals(result.window().topLeftAddress()),
            "window topLeftAddress mismatch");
        require(selector.rowCount() == result.window().rowCount(), "window rowCount mismatch");
        require(
            selector.columnCount() == result.window().columnCount(), "window columnCount mismatch");
      }
      case InspectionQuery.GetMergedRegions _ -> {
        InspectionResult.MergedRegionsResult result =
            (InspectionResult.MergedRegionsResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target()).name().equals(result.sheetName()),
            "merged regions sheet mismatch");
      }
      case InspectionQuery.GetHyperlinks _ -> {
        InspectionResult.HyperlinksResult result = (InspectionResult.HyperlinksResult) readResult;
        require(
            sheetName((CellSelector) readOperation.target()).equals(result.sheetName()),
            "hyperlinks sheet mismatch");
      }
      case InspectionQuery.GetComments _ -> {
        InspectionResult.CommentsResult result = (InspectionResult.CommentsResult) readResult;
        require(
            sheetName((CellSelector) readOperation.target()).equals(result.sheetName()),
            "comments sheet mismatch");
      }
      case InspectionQuery.GetDrawingObjects _ -> {
        InspectionResult.DrawingObjectsResult result =
            (InspectionResult.DrawingObjectsResult) readResult;
        require(
            ((DrawingObjectSelector.AllOnSheet) readOperation.target())
                .sheetName()
                .equals(result.sheetName()),
            "drawing objects sheet mismatch");
        result.drawingObjects().forEach(WorkbookInvariantResponseChecks::requireDrawingObjectShape);
      }
      case InspectionQuery.GetCharts _ -> {
        InspectionResult.ChartsResult result = (InspectionResult.ChartsResult) readResult;
        require(
            ((ChartSelector.AllOnSheet) readOperation.target())
                .sheetName()
                .equals(result.sheetName()),
            "charts sheet mismatch");
        result.charts().forEach(WorkbookInvariantResponseChecks::requireChartReportShape);
      }
      case InspectionQuery.GetPivotTables _ -> {
        InspectionResult.PivotTablesResult result = (InspectionResult.PivotTablesResult) readResult;
        result.pivotTables().forEach(WorkbookInvariantResponseChecks::requirePivotTableShape);
      }
      case InspectionQuery.GetDrawingObjectPayload _ -> {
        InspectionResult.DrawingObjectPayloadResult result =
            (InspectionResult.DrawingObjectPayloadResult) readResult;
        DrawingObjectSelector.ByName selector =
            (DrawingObjectSelector.ByName) readOperation.target();
        require(selector.sheetName().equals(result.sheetName()), "drawing payload sheet mismatch");
        WorkbookInvariantResponseChecks.requireDrawingObjectPayloadShape(result.payload());
        require(
            selector.objectName().equals(result.payload().name()),
            "drawing payload objectName mismatch");
      }
      case InspectionQuery.GetSheetLayout _ -> {
        InspectionResult.SheetLayoutResult result = (InspectionResult.SheetLayoutResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target())
                .name()
                .equals(result.layout().sheetName()),
            "layout sheet mismatch");
      }
      case InspectionQuery.GetPrintLayout _ -> {
        InspectionResult.PrintLayoutResult result = (InspectionResult.PrintLayoutResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target())
                .name()
                .equals(result.layout().sheetName()),
            "print layout sheet mismatch");
      }
      case InspectionQuery.GetDataValidations _ -> {
        InspectionResult.DataValidationsResult result =
            (InspectionResult.DataValidationsResult) readResult;
        require(
            sheetName((RangeSelector) readOperation.target()).equals(result.sheetName()),
            "data validations sheet mismatch");
      }
      case InspectionQuery.GetConditionalFormatting _ -> {
        InspectionResult.ConditionalFormattingResult result =
            (InspectionResult.ConditionalFormattingResult) readResult;
        require(
            sheetName((RangeSelector) readOperation.target()).equals(result.sheetName()),
            "conditional formatting sheet mismatch");
      }
      case InspectionQuery.GetAutofilters _ -> {
        InspectionResult.AutofiltersResult result = (InspectionResult.AutofiltersResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target()).name().equals(result.sheetName()),
            "autofilters sheet mismatch");
      }
      case InspectionQuery.GetTables _ -> {
        InspectionResult.TablesResult result = (InspectionResult.TablesResult) readResult;
        result.tables().forEach(WorkbookInvariantResponseChecks::requireTableEntryShape);
      }
      case InspectionQuery.GetFormulaSurface _ -> {
        InspectionResult.FormulaSurfaceResult result =
            (InspectionResult.FormulaSurfaceResult) readResult;
        require(result.analysis().sheets() != null, "formula surface sheets must not be null");
      }
      case InspectionQuery.GetSheetSchema _ -> {
        InspectionResult.SheetSchemaResult result = (InspectionResult.SheetSchemaResult) readResult;
        RangeSelector.RectangularWindow selector =
            (RangeSelector.RectangularWindow) readOperation.target();
        require(
            selector.sheetName().equals(result.analysis().sheetName()), "schema sheet mismatch");
        require(
            selector.topLeftAddress().equals(result.analysis().topLeftAddress()),
            "schema topLeftAddress mismatch");
      }
      case InspectionQuery.GetNamedRangeSurface _ -> {
        InspectionResult.NamedRangeSurfaceResult result =
            (InspectionResult.NamedRangeSurfaceResult) readResult;
        require(
            result.analysis().namedRanges() != null,
            "named range surface entries must not be null");
      }
      case InspectionQuery.AnalyzeFormulaHealth _ ->
          WorkbookInvariantResponseChecks.requireFormulaHealthShape(
              ((InspectionResult.FormulaHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeDataValidationHealth _ ->
          WorkbookInvariantResponseChecks.requireDataValidationHealthShape(
              ((InspectionResult.DataValidationHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeConditionalFormattingHealth _ ->
          WorkbookInvariantResponseChecks.requireConditionalFormattingHealthShape(
              ((InspectionResult.ConditionalFormattingHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeAutofilterHealth _ ->
          WorkbookInvariantResponseChecks.requireAutofilterHealthShape(
              ((InspectionResult.AutofilterHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeTableHealth _ ->
          WorkbookInvariantResponseChecks.requireTableHealthShape(
              ((InspectionResult.TableHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzePivotTableHealth _ ->
          WorkbookInvariantResponseChecks.requirePivotTableHealthShape(
              ((InspectionResult.PivotTableHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeHyperlinkHealth _ ->
          WorkbookInvariantResponseChecks.requireHyperlinkHealthShape(
              ((InspectionResult.HyperlinkHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeNamedRangeHealth _ ->
          WorkbookInvariantResponseChecks.requireNamedRangeHealthShape(
              ((InspectionResult.NamedRangeHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeWorkbookFindings _ ->
          WorkbookInvariantResponseChecks.requireWorkbookFindingsShape(
              ((InspectionResult.WorkbookFindingsResult) readResult).analysis());
    }
  }

  private static String sheetName(CellSelector selector) {
    return switch (selector) {
      case CellSelector.AllUsedInSheet all -> all.sheetName();
      case CellSelector.ByAddress byAddress -> byAddress.sheetName();
      case CellSelector.ByAddresses byAddresses -> byAddresses.sheetName();
      case CellSelector.ByQualifiedAddresses _ -> null;
    };
  }

  private static String sheetName(RangeSelector selector) {
    return switch (selector) {
      case RangeSelector.AllOnSheet allOnSheet -> allOnSheet.sheetName();
      case RangeSelector.ByRange byRange -> byRange.sheetName();
      case RangeSelector.ByRanges byRanges -> byRanges.sheetName();
      case RangeSelector.RectangularWindow window -> window.sheetName();
    };
  }

  private static String readResultKind(InspectionResult readResult) {
    return switch (readResult) {
      case InspectionResult.WorkbookSummaryResult _ -> "GET_WORKBOOK_SUMMARY";
      case InspectionResult.PackageSecurityResult _ -> "GET_PACKAGE_SECURITY";
      case InspectionResult.WorkbookProtectionResult _ -> "GET_WORKBOOK_PROTECTION";
      case InspectionResult.CustomXmlMappingsResult _ -> "GET_CUSTOM_XML_MAPPINGS";
      case InspectionResult.CustomXmlExportResult _ -> "EXPORT_CUSTOM_XML_MAPPING";
      case InspectionResult.NamedRangesResult _ -> "GET_NAMED_RANGES";
      case InspectionResult.SheetSummaryResult _ -> "GET_SHEET_SUMMARY";
      case InspectionResult.ArrayFormulasResult _ -> "GET_ARRAY_FORMULAS";
      case InspectionResult.CellsResult _ -> "GET_CELLS";
      case InspectionResult.WindowResult _ -> "GET_WINDOW";
      case InspectionResult.MergedRegionsResult _ -> "GET_MERGED_REGIONS";
      case InspectionResult.HyperlinksResult _ -> "GET_HYPERLINKS";
      case InspectionResult.CommentsResult _ -> "GET_COMMENTS";
      case InspectionResult.DrawingObjectsResult _ -> "GET_DRAWING_OBJECTS";
      case InspectionResult.ChartsResult _ -> "GET_CHARTS";
      case InspectionResult.PivotTablesResult _ -> "GET_PIVOT_TABLES";
      case InspectionResult.DrawingObjectPayloadResult _ -> "GET_DRAWING_OBJECT_PAYLOAD";
      case InspectionResult.SheetLayoutResult _ -> "GET_SHEET_LAYOUT";
      case InspectionResult.PrintLayoutResult _ -> "GET_PRINT_LAYOUT";
      case InspectionResult.DataValidationsResult _ -> "GET_DATA_VALIDATIONS";
      case InspectionResult.ConditionalFormattingResult _ -> "GET_CONDITIONAL_FORMATTING";
      case InspectionResult.AutofiltersResult _ -> "GET_AUTOFILTERS";
      case InspectionResult.TablesResult _ -> "GET_TABLES";
      case InspectionResult.FormulaSurfaceResult _ -> "GET_FORMULA_SURFACE";
      case InspectionResult.SheetSchemaResult _ -> "GET_SHEET_SCHEMA";
      case InspectionResult.NamedRangeSurfaceResult _ -> "GET_NAMED_RANGE_SURFACE";
      case InspectionResult.FormulaHealthResult _ -> "ANALYZE_FORMULA_HEALTH";
      case InspectionResult.DataValidationHealthResult _ -> "ANALYZE_DATA_VALIDATION_HEALTH";
      case InspectionResult.ConditionalFormattingHealthResult _ ->
          "ANALYZE_CONDITIONAL_FORMATTING_HEALTH";
      case InspectionResult.AutofilterHealthResult _ -> "ANALYZE_AUTOFILTER_HEALTH";
      case InspectionResult.TableHealthResult _ -> "ANALYZE_TABLE_HEALTH";
      case InspectionResult.PivotTableHealthResult _ -> "ANALYZE_PIVOT_TABLE_HEALTH";
      case InspectionResult.HyperlinkHealthResult _ -> "ANALYZE_HYPERLINK_HEALTH";
      case InspectionResult.NamedRangeHealthResult _ -> "ANALYZE_NAMED_RANGE_HEALTH";
      case InspectionResult.WorkbookFindingsResult _ -> "ANALYZE_WORKBOOK_FINDINGS";
    };
  }
}
