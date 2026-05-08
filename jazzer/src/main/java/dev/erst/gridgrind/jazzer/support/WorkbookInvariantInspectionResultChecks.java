package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks.require;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAnalysisResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookSurfaceInspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import java.util.Objects;
import java.util.Optional;

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
      case WorkbookIntrospectionQuery.GetWorkbookSummary _ -> {
        WorkbookInspectionResult.WorkbookSummaryResult result =
            (WorkbookInspectionResult.WorkbookSummaryResult) readResult;
        WorkbookInvariantResponseChecks.requireWorkbookSummaryShape(result.workbook());
      }
      case WorkbookIntrospectionQuery.GetPackageSecurity _ ->
          WorkbookInvariantResponseChecks.requirePackageSecurityShape(
              ((WorkbookInspectionResult.PackageSecurityResult) readResult).security());
      case WorkbookIntrospectionQuery.GetWorkbookProtection _ ->
          WorkbookInvariantResponseChecks.requireWorkbookProtectionShape(
              ((WorkbookInspectionResult.WorkbookProtectionResult) readResult).protection());
      case WorkbookIntrospectionQuery.GetCustomXmlMappings _ -> {
        WorkbookInspectionResult.CustomXmlMappingsResult result =
            (WorkbookInspectionResult.CustomXmlMappingsResult) readResult;
        result.mappings().forEach(WorkbookInvariantResponseChecks::requireCustomXmlMappingShape);
      }
      case WorkbookIntrospectionQuery.ExportCustomXmlMapping _ ->
          WorkbookInvariantResponseChecks.requireCustomXmlExportShape(
              ((WorkbookInspectionResult.CustomXmlExportResult) readResult).export());
      case WorkbookIntrospectionQuery.GetNamedRanges _ -> {
        WorkbookInspectionResult.NamedRangesResult result =
            (WorkbookInspectionResult.NamedRangesResult) readResult;
        result.namedRanges().forEach(WorkbookInvariantResponseChecks::requireNamedRangeShape);
      }
      case SheetIntrospectionQuery.GetSheetSummary _ -> {
        SheetInspectionResult.SheetSummaryResult result =
            (SheetInspectionResult.SheetSummaryResult) readResult;
        WorkbookInvariantResponseChecks.requireSheetSummaryShape(result.sheet());
        require(
            ((SheetSelector.ByName) readOperation.target())
                .name()
                .equals(result.sheet().sheetName()),
            "sheet summary sheet mismatch");
      }
      case SheetIntrospectionQuery.GetArrayFormulas _ -> {
        SheetInspectionResult.ArrayFormulasResult result =
            (SheetInspectionResult.ArrayFormulasResult) readResult;
        result.arrayFormulas().forEach(WorkbookInvariantResponseChecks::requireArrayFormulaShape);
      }
      case SheetIntrospectionQuery.GetCells _ -> {
        SheetInspectionResult.CellsResult result = (SheetInspectionResult.CellsResult) readResult;
        require(
            Objects.equals(
                sheetName((CellSelector) readOperation.target()).orElse(null), result.sheetName()),
            "cells sheet mismatch");
        if (readOperation.target() instanceof CellSelector.ByAddresses byAddresses) {
          require(
              result.cells().size() == byAddresses.addresses().size(),
              "cells result size must match requested addresses");
        } else if (readOperation.target() instanceof CellSelector.ByAddress) {
          require(result.cells().size() == 1, "single-cell result size must be 1");
        }
      }
      case SheetIntrospectionQuery.GetWindow _ -> {
        SheetInspectionResult.WindowResult result = (SheetInspectionResult.WindowResult) readResult;
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
      case SheetIntrospectionQuery.GetMergedRegions _ -> {
        SheetInspectionResult.MergedRegionsResult result =
            (SheetInspectionResult.MergedRegionsResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target()).name().equals(result.sheetName()),
            "merged regions sheet mismatch");
      }
      case SheetIntrospectionQuery.GetHyperlinks _ -> {
        SheetInspectionResult.HyperlinksResult result =
            (SheetInspectionResult.HyperlinksResult) readResult;
        require(
            Objects.equals(
                sheetName((CellSelector) readOperation.target()).orElse(null), result.sheetName()),
            "hyperlinks sheet mismatch");
      }
      case SheetIntrospectionQuery.GetComments _ -> {
        SheetInspectionResult.CommentsResult result =
            (SheetInspectionResult.CommentsResult) readResult;
        require(
            Objects.equals(
                sheetName((CellSelector) readOperation.target()).orElse(null), result.sheetName()),
            "comments sheet mismatch");
      }
      case WorkbookAssetIntrospectionQuery.GetDrawingObjects _ -> {
        WorkbookAssetInspectionResult.DrawingObjectsResult result =
            (WorkbookAssetInspectionResult.DrawingObjectsResult) readResult;
        require(
            ((DrawingObjectSelector.AllOnSheet) readOperation.target())
                .sheetName()
                .equals(result.sheetName()),
            "drawing objects sheet mismatch");
        result.drawingObjects().forEach(WorkbookInvariantResponseChecks::requireDrawingObjectShape);
      }
      case WorkbookAssetIntrospectionQuery.GetCharts _ -> {
        WorkbookAssetInspectionResult.ChartsResult result =
            (WorkbookAssetInspectionResult.ChartsResult) readResult;
        require(
            ((ChartSelector.AllOnSheet) readOperation.target())
                .sheetName()
                .equals(result.sheetName()),
            "charts sheet mismatch");
        result.charts().forEach(WorkbookInvariantResponseChecks::requireChartReportShape);
      }
      case WorkbookAssetIntrospectionQuery.GetPivotTables _ -> {
        WorkbookAssetInspectionResult.PivotTablesResult result =
            (WorkbookAssetInspectionResult.PivotTablesResult) readResult;
        result.pivotTables().forEach(WorkbookInvariantResponseChecks::requirePivotTableShape);
      }
      case WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload _ -> {
        WorkbookAssetInspectionResult.DrawingObjectPayloadResult result =
            (WorkbookAssetInspectionResult.DrawingObjectPayloadResult) readResult;
        DrawingObjectSelector.ByName selector =
            (DrawingObjectSelector.ByName) readOperation.target();
        require(selector.sheetName().equals(result.sheetName()), "drawing payload sheet mismatch");
        WorkbookInvariantResponseChecks.requireDrawingObjectPayloadShape(result.payload());
        require(
            selector.objectName().equals(result.payload().name()),
            "drawing payload objectName mismatch");
      }
      case SheetIntrospectionQuery.GetSheetLayout _ -> {
        SheetInspectionResult.SheetLayoutResult result =
            (SheetInspectionResult.SheetLayoutResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target())
                .name()
                .equals(result.layout().sheetName()),
            "layout sheet mismatch");
      }
      case SheetIntrospectionQuery.GetPrintLayout _ -> {
        SheetInspectionResult.PrintLayoutResult result =
            (SheetInspectionResult.PrintLayoutResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target())
                .name()
                .equals(result.layout().sheetName()),
            "print layout sheet mismatch");
      }
      case SheetIntrospectionQuery.GetDataValidations _ -> {
        SheetInspectionResult.DataValidationsResult result =
            (SheetInspectionResult.DataValidationsResult) readResult;
        require(
            sheetName((RangeSelector) readOperation.target()).equals(result.sheetName()),
            "data validations sheet mismatch");
      }
      case SheetIntrospectionQuery.GetConditionalFormatting _ -> {
        SheetInspectionResult.ConditionalFormattingResult result =
            (SheetInspectionResult.ConditionalFormattingResult) readResult;
        require(
            sheetName((RangeSelector) readOperation.target()).equals(result.sheetName()),
            "conditional formatting sheet mismatch");
      }
      case SheetIntrospectionQuery.GetAutofilters _ -> {
        SheetInspectionResult.AutofiltersResult result =
            (SheetInspectionResult.AutofiltersResult) readResult;
        require(
            ((SheetSelector.ByName) readOperation.target()).name().equals(result.sheetName()),
            "autofilters sheet mismatch");
      }
      case WorkbookAssetIntrospectionQuery.GetTables _ -> {
        WorkbookAssetInspectionResult.TablesResult result =
            (WorkbookAssetInspectionResult.TablesResult) readResult;
        result.tables().forEach(WorkbookInvariantResponseChecks::requireTableEntryShape);
      }
      case InspectionSurfaceQuery.GetFormulaSurface _ -> {
        WorkbookSurfaceInspectionResult.FormulaSurfaceResult result =
            (WorkbookSurfaceInspectionResult.FormulaSurfaceResult) readResult;
        require(result.surface().sheets() != null, "formula surface sheets must not be null");
      }
      case InspectionSurfaceQuery.GetSheetSchema _ -> {
        WorkbookSurfaceInspectionResult.SheetSchemaResult result =
            (WorkbookSurfaceInspectionResult.SheetSchemaResult) readResult;
        RangeSelector.RectangularWindow selector =
            (RangeSelector.RectangularWindow) readOperation.target();
        require(selector.sheetName().equals(result.surface().sheetName()), "schema sheet mismatch");
        require(
            selector.topLeftAddress().equals(result.surface().topLeftAddress()),
            "schema topLeftAddress mismatch");
      }
      case InspectionSurfaceQuery.GetNamedRangeSurface _ -> {
        WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult result =
            (WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult) readResult;
        require(
            result.surface().namedRanges() != null, "named range surface entries must not be null");
      }
      case InspectionAnalysisQuery.AnalyzeFormulaHealth _ ->
          WorkbookInvariantResponseChecks.requireFormulaHealthShape(
              ((WorkbookAnalysisResult.FormulaHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzeDataValidationHealth _ ->
          WorkbookInvariantResponseChecks.requireDataValidationHealthShape(
              ((WorkbookAnalysisResult.DataValidationHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth _ ->
          WorkbookInvariantResponseChecks.requireConditionalFormattingHealthShape(
              ((WorkbookAnalysisResult.ConditionalFormattingHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzeAutofilterHealth _ ->
          WorkbookInvariantResponseChecks.requireAutofilterHealthShape(
              ((WorkbookAnalysisResult.AutofilterHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzeTableHealth _ ->
          WorkbookInvariantResponseChecks.requireTableHealthShape(
              ((WorkbookAnalysisResult.TableHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzePivotTableHealth _ ->
          WorkbookInvariantResponseChecks.requirePivotTableHealthShape(
              ((WorkbookAnalysisResult.PivotTableHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzeHyperlinkHealth _ ->
          WorkbookInvariantResponseChecks.requireHyperlinkHealthShape(
              ((WorkbookAnalysisResult.HyperlinkHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzeNamedRangeHealth _ ->
          WorkbookInvariantResponseChecks.requireNamedRangeHealthShape(
              ((WorkbookAnalysisResult.NamedRangeHealthResult) readResult).analysis());
      case InspectionAnalysisQuery.AnalyzeWorkbookFindings _ ->
          WorkbookInvariantResponseChecks.requireWorkbookFindingsShape(
              ((WorkbookAnalysisResult.WorkbookFindingsResult) readResult).analysis());
    }
  }

  private static Optional<String> sheetName(CellSelector selector) {
    return switch (selector) {
      case CellSelector.AllUsedInSheet all -> Optional.of(all.sheetName());
      case CellSelector.ByAddress byAddress -> Optional.of(byAddress.sheetName());
      case CellSelector.ByAddresses byAddresses -> Optional.of(byAddresses.sheetName());
      case CellSelector.ByQualifiedAddresses _ -> Optional.empty();
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
      case WorkbookInspectionResult.WorkbookSummaryResult _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookInspectionResult.PackageSecurityResult _ -> "GET_PACKAGE_SECURITY";
      case WorkbookInspectionResult.WorkbookProtectionResult _ -> "GET_WORKBOOK_PROTECTION";
      case WorkbookInspectionResult.CustomXmlMappingsResult _ -> "GET_CUSTOM_XML_MAPPINGS";
      case WorkbookInspectionResult.CustomXmlExportResult _ -> "EXPORT_CUSTOM_XML_MAPPING";
      case WorkbookInspectionResult.NamedRangesResult _ -> "GET_NAMED_RANGES";
      case SheetInspectionResult.SheetSummaryResult _ -> "GET_SHEET_SUMMARY";
      case SheetInspectionResult.ArrayFormulasResult _ -> "GET_ARRAY_FORMULAS";
      case SheetInspectionResult.CellsResult _ -> "GET_CELLS";
      case SheetInspectionResult.WindowResult _ -> "GET_WINDOW";
      case SheetInspectionResult.MergedRegionsResult _ -> "GET_MERGED_REGIONS";
      case SheetInspectionResult.HyperlinksResult _ -> "GET_HYPERLINKS";
      case SheetInspectionResult.CommentsResult _ -> "GET_COMMENTS";
      case WorkbookAssetInspectionResult.DrawingObjectsResult _ -> "GET_DRAWING_OBJECTS";
      case WorkbookAssetInspectionResult.ChartsResult _ -> "GET_CHARTS";
      case WorkbookAssetInspectionResult.PivotTablesResult _ -> "GET_PIVOT_TABLES";
      case WorkbookAssetInspectionResult.DrawingObjectPayloadResult _ ->
          "GET_DRAWING_OBJECT_PAYLOAD";
      case SheetInspectionResult.SheetLayoutResult _ -> "GET_SHEET_LAYOUT";
      case SheetInspectionResult.PrintLayoutResult _ -> "GET_PRINT_LAYOUT";
      case SheetInspectionResult.DataValidationsResult _ -> "GET_DATA_VALIDATIONS";
      case SheetInspectionResult.ConditionalFormattingResult _ -> "GET_CONDITIONAL_FORMATTING";
      case SheetInspectionResult.AutofiltersResult _ -> "GET_AUTOFILTERS";
      case WorkbookAssetInspectionResult.TablesResult _ -> "GET_TABLES";
      case WorkbookSurfaceInspectionResult.FormulaSurfaceResult _ -> "GET_FORMULA_SURFACE";
      case WorkbookSurfaceInspectionResult.SheetSchemaResult _ -> "GET_SHEET_SCHEMA";
      case WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult _ -> "GET_NAMED_RANGE_SURFACE";
      case WorkbookAnalysisResult.FormulaHealthResult _ -> "ANALYZE_FORMULA_HEALTH";
      case WorkbookAnalysisResult.DataValidationHealthResult _ -> "ANALYZE_DATA_VALIDATION_HEALTH";
      case WorkbookAnalysisResult.ConditionalFormattingHealthResult _ ->
          "ANALYZE_CONDITIONAL_FORMATTING_HEALTH";
      case WorkbookAnalysisResult.AutofilterHealthResult _ -> "ANALYZE_AUTOFILTER_HEALTH";
      case WorkbookAnalysisResult.TableHealthResult _ -> "ANALYZE_TABLE_HEALTH";
      case WorkbookAnalysisResult.PivotTableHealthResult _ -> "ANALYZE_PIVOT_TABLE_HEALTH";
      case WorkbookAnalysisResult.HyperlinkHealthResult _ -> "ANALYZE_HYPERLINK_HEALTH";
      case WorkbookAnalysisResult.NamedRangeHealthResult _ -> "ANALYZE_NAMED_RANGE_HEALTH";
      case WorkbookAnalysisResult.WorkbookFindingsResult _ -> "ANALYZE_WORKBOOK_FINDINGS";
    };
  }
}
