package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.Objects;

/** Validates protocol responses and workbook state without depending on JUnit assertions. */
public final class WorkbookInvariantChecks {
  private WorkbookInvariantChecks() {}

  /** Requires the response shape to satisfy the protocol invariants the fuzzers rely on. */
  public static void requireResponseShape(GridGrindResponse response) {
    require(response != null, "response must not be null");
    require(response.protocolVersion() != null, "protocolVersion must not be null");

    switch (response) {
      case GridGrindResponse.Success success -> requireSuccessResponseShape(success);
      case GridGrindResponse.Failure failure -> requireFailureResponseShape(failure);
    }
  }

  /**
   * Requires the response shape to agree with the request's source, reads, and persistence
   * contract.
   */
  public static void requireWorkflowOutcomeShape(WorkbookPlan request, GridGrindResponse response) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(response, "response must not be null");

    switch (response) {
      case GridGrindResponse.Failure _ -> {
        return;
      }
      case GridGrindResponse.Success success -> {
        requirePersistenceMatchesRequest(request.persistence(), success.persistence());
        require(
            success.assertions().size() == request.assertionSteps().size(),
            "assertions size must match the requested assertion count");
        for (int index = 0; index < request.assertionSteps().size(); index++) {
          requireAssertionMatchesRequest(
              request.assertionSteps().get(index), success.assertions().get(index));
        }
        require(
            success.inspections().size() == request.inspectionSteps().size(),
            "inspections size must match the requested inspection count");
        for (int index = 0; index < request.inspectionSteps().size(); index++) {
          requireReadMatchesRequest(
              request.inspectionSteps().get(index), success.inspections().get(index));
        }
      }
    }
  }

  /** Requires the open workbook to satisfy the structural invariants the fuzzers rely on. */
  public static void requireWorkbookShape(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    var workbookSummary =
        ((dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult)
                readExecutor
                    .apply(workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook-shape"))
                    .getFirst())
            .workbook();

    requireEngineWorkbookSummaryShape(workbookSummary);
    workbook
        .sheetNames()
        .forEach(
            sheetName -> {
              requireEngineSheetSummaryShape(
                  ((dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult)
                          readExecutor
                              .apply(
                                  workbook,
                                  new WorkbookReadCommand.GetSheetSummary(
                                      "sheet-shape-" + sheetName, sheetName))
                              .getFirst())
                      .sheet());
              ((dev.erst.gridgrind.excel.WorkbookReadResult.DrawingObjectsResult)
                      readExecutor
                          .apply(
                              workbook,
                              new WorkbookReadCommand.GetDrawingObjects(
                                  "drawing-shape-" + sheetName, sheetName))
                          .getFirst())
                  .drawingObjects()
                  .forEach(WorkbookInvariantChecks::requireEngineDrawingObjectShape);
              ((dev.erst.gridgrind.excel.WorkbookReadResult.ChartsResult)
                      readExecutor
                          .apply(
                              workbook,
                              new WorkbookReadCommand.GetCharts(
                                  "chart-shape-" + sheetName, sheetName))
                          .getFirst())
                  .charts()
                  .forEach(WorkbookInvariantChecks::requireEngineChartShape);
            });
    ((dev.erst.gridgrind.excel.WorkbookReadResult.PivotTablesResult)
            readExecutor
                .apply(
                    workbook,
                    new WorkbookReadCommand.GetPivotTables(
                        "pivot-shape", new dev.erst.gridgrind.excel.ExcelPivotTableSelection.All()))
                .getFirst())
        .pivotTables()
        .forEach(WorkbookInvariantChecks::requireEnginePivotTableShape);
  }

  private static void requireSuccessResponseShape(GridGrindResponse.Success success) {
    require(success.persistence() != null, "persistence must not be null");
    requirePersistenceOutcomeShape(success.persistence());
    require(success.warnings() != null, "warnings must not be null");
    success.warnings().forEach(WorkbookInvariantChecks::requireRequestWarningShape);
    require(success.assertions() != null, "assertions must not be null");
    success.assertions().forEach(WorkbookInvariantChecks::requireAssertionResultShape);
    require(success.inspections() != null, "inspections must not be null");
    success.inspections().forEach(WorkbookInvariantChecks::requireReadResultShape);
  }

  private static void requireFailureResponseShape(GridGrindResponse.Failure failure) {
    require(failure.problem() != null, "problem must not be null");
    require(failure.problem().code() != null, "problem code must not be null");
    require(failure.problem().category() != null, "problem category must not be null");
    require(failure.problem().recovery() != null, "problem recovery must not be null");
    require(failure.problem().title() != null, "problem title must not be null");
    require(failure.problem().message() != null, "problem message must not be null");
    require(failure.problem().resolution() != null, "problem resolution must not be null");
    require(failure.problem().context() != null, "problem context must not be null");
    require(failure.problem().causes() != null, "problem causes must not be null");
    if (failure.problem().assertionFailure().isPresent()) {
      requireAssertionFailureShape(failure.problem().assertionFailure().orElseThrow());
      require(
          failure.problem().code()
              == dev.erst.gridgrind.contract.dto.GridGrindProblemCode.ASSERTION_FAILED,
          "only ASSERTION_FAILED problems may carry assertionFailure details");
      return;
    }
    require(
        failure.problem().code()
            != dev.erst.gridgrind.contract.dto.GridGrindProblemCode.ASSERTION_FAILED,
        "ASSERTION_FAILED problems must carry assertionFailure details");
  }

  private static void requirePersistenceMatchesRequest(
      WorkbookPlan.WorkbookPersistence requestPersistence,
      GridGrindResponse.PersistenceOutcome persistenceOutcome) {
    switch (requestPersistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> {
        switch (persistenceOutcome) {
          case GridGrindResponse.PersistenceOutcome.NotSaved _ -> {}
          case GridGrindResponse.PersistenceOutcome.SavedAs _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
          case GridGrindResponse.PersistenceOutcome.Overwritten _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
        }
      }
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> {
        requirePersistenceOutcomeShape(persistenceOutcome);
        require(
            persistenceOutcome instanceof GridGrindResponse.PersistenceOutcome.Overwritten,
            "OVERWRITE persistence must return OVERWRITE outcome");
      }
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> {
        requirePersistenceOutcomeShape(persistenceOutcome);
        require(
            persistenceOutcome instanceof GridGrindResponse.PersistenceOutcome.SavedAs,
            "SAVE_AS persistence must return SAVE_AS outcome");
      }
    }
  }

  private static void requirePersistenceOutcomeShape(
      GridGrindResponse.PersistenceOutcome persistenceOutcome) {
    switch (persistenceOutcome) {
      case GridGrindResponse.PersistenceOutcome.NotSaved _ -> {}
      case GridGrindResponse.PersistenceOutcome.SavedAs savedAs -> {
        requireNonBlank(savedAs.requestedPath(), "requestedPath");
        requireExecutionWorkbookPath(savedAs.executionPath());
      }
      case GridGrindResponse.PersistenceOutcome.Overwritten overwritten -> {
        requireNonBlank(overwritten.sourcePath(), "sourcePath");
        requireExecutionWorkbookPath(overwritten.executionPath());
      }
    }
  }

  private static void requireExecutionWorkbookPath(String executionPath) {
    require(executionPath != null, "executionPath must not be null");
    require(executionPath.endsWith(".xlsx"), "executionPath must point to .xlsx");
    require(Files.exists(Path.of(executionPath)), "executionPath must exist");
  }

  private static void requireRequestWarningShape(RequestWarning warning) {
    require(warning != null, "warning must not be null");
    require(warning.stepIndex() >= 0, "warning stepIndex must not be negative");
    requireNonBlank(warning.stepType(), "warning stepType");
    requireNonBlank(warning.message(), "warning message");
  }

  private static void requireReadMatchesRequest(
      InspectionStep readOperation, InspectionResult readResult) {
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
        requireWorkbookSummaryShape(result.workbook());
      }
      case InspectionQuery.GetPackageSecurity _ ->
          requirePackageSecurityShape(
              ((InspectionResult.PackageSecurityResult) readResult).security());
      case InspectionQuery.GetWorkbookProtection _ ->
          requireWorkbookProtectionShape(
              ((InspectionResult.WorkbookProtectionResult) readResult).protection());
      case InspectionQuery.GetCustomXmlMappings _ -> {
        InspectionResult.CustomXmlMappingsResult result =
            (InspectionResult.CustomXmlMappingsResult) readResult;
        result.mappings().forEach(WorkbookInvariantChecks::requireCustomXmlMappingShape);
      }
      case InspectionQuery.ExportCustomXmlMapping _ ->
          requireCustomXmlExportShape(
              ((InspectionResult.CustomXmlExportResult) readResult).export());
      case InspectionQuery.GetNamedRanges _ -> {
        InspectionResult.NamedRangesResult result = (InspectionResult.NamedRangesResult) readResult;
        result.namedRanges().forEach(WorkbookInvariantChecks::requireNamedRangeShape);
      }
      case InspectionQuery.GetSheetSummary _ -> {
        InspectionResult.SheetSummaryResult result =
            (InspectionResult.SheetSummaryResult) readResult;
        requireSheetSummaryShape(result.sheet());
        require(
            ((SheetSelector.ByName) readOperation.target())
                .name()
                .equals(result.sheet().sheetName()),
            "sheet summary sheet mismatch");
      }
      case InspectionQuery.GetArrayFormulas _ -> {
        InspectionResult.ArrayFormulasResult result =
            (InspectionResult.ArrayFormulasResult) readResult;
        result.arrayFormulas().forEach(WorkbookInvariantChecks::requireArrayFormulaShape);
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
        result.drawingObjects().forEach(WorkbookInvariantChecks::requireDrawingObjectShape);
      }
      case InspectionQuery.GetCharts _ -> {
        InspectionResult.ChartsResult result = (InspectionResult.ChartsResult) readResult;
        require(
            ((ChartSelector.AllOnSheet) readOperation.target())
                .sheetName()
                .equals(result.sheetName()),
            "charts sheet mismatch");
        result.charts().forEach(WorkbookInvariantChecks::requireChartReportShape);
      }
      case InspectionQuery.GetPivotTables _ -> {
        InspectionResult.PivotTablesResult result = (InspectionResult.PivotTablesResult) readResult;
        result.pivotTables().forEach(WorkbookInvariantChecks::requirePivotTableShape);
      }
      case InspectionQuery.GetDrawingObjectPayload _ -> {
        InspectionResult.DrawingObjectPayloadResult result =
            (InspectionResult.DrawingObjectPayloadResult) readResult;
        DrawingObjectSelector.ByName selector =
            (DrawingObjectSelector.ByName) readOperation.target();
        require(selector.sheetName().equals(result.sheetName()), "drawing payload sheet mismatch");
        requireDrawingObjectPayloadShape(result.payload());
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
        result.tables().forEach(WorkbookInvariantChecks::requireTableEntryShape);
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
          requireFormulaHealthShape(((InspectionResult.FormulaHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeDataValidationHealth _ ->
          requireDataValidationHealthShape(
              ((InspectionResult.DataValidationHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeConditionalFormattingHealth _ ->
          requireConditionalFormattingHealthShape(
              ((InspectionResult.ConditionalFormattingHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeAutofilterHealth _ ->
          requireAutofilterHealthShape(
              ((InspectionResult.AutofilterHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeTableHealth _ ->
          requireTableHealthShape(((InspectionResult.TableHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzePivotTableHealth _ ->
          requirePivotTableHealthShape(
              ((InspectionResult.PivotTableHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeHyperlinkHealth _ ->
          requireHyperlinkHealthShape(
              ((InspectionResult.HyperlinkHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeNamedRangeHealth _ ->
          requireNamedRangeHealthShape(
              ((InspectionResult.NamedRangeHealthResult) readResult).analysis());
      case InspectionQuery.AnalyzeWorkbookFindings _ ->
          requireWorkbookFindingsShape(
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

  private static void requireAssertionMatchesRequest(
      dev.erst.gridgrind.contract.step.AssertionStep assertionStep,
      AssertionResult assertionResult) {
    require(
        assertionStep.stepId().equals(assertionResult.stepId()),
        "assertion result stepId must match the request");
    require(
        SequenceIntrospection.assertionKind(assertionStep).equals(assertionResult.assertionType()),
        "assertion result type must match the requested assertion kind");
  }

  private static void requireAssertionResultShape(AssertionResult assertionResult) {
    require(assertionResult != null, "assertion result must not be null");
    requireNonBlank(assertionResult.stepId(), "assertion stepId");
    requireNonBlank(assertionResult.assertionType(), "assertionType");
  }

  private static void requireAssertionFailureShape(AssertionFailure assertionFailure) {
    require(assertionFailure != null, "assertionFailure must not be null");
    requireNonBlank(assertionFailure.stepId(), "assertionFailure stepId");
    requireNonBlank(assertionFailure.assertionType(), "assertionFailure assertionType");
    require(assertionFailure.target() != null, "assertionFailure target must not be null");
    require(assertionFailure.assertion() != null, "assertionFailure assertion must not be null");
    require(
        assertionFailure.observations() != null, "assertionFailure observations must not be null");
    assertionFailure.observations().forEach(WorkbookInvariantChecks::requireReadResultShape);
  }

  private static void requireReadResultShape(InspectionResult readResult) {
    require(readResult.stepId() != null, "read stepId must not be null");
    require(!readResult.stepId().isBlank(), "read stepId must not be blank");

    switch (readResult) {
      case InspectionResult.WorkbookSummaryResult result ->
          requireWorkbookSummaryShape(result.workbook());
      case InspectionResult.PackageSecurityResult result ->
          requirePackageSecurityShape(result.security());
      case InspectionResult.WorkbookProtectionResult result ->
          requireWorkbookProtectionShape(result.protection());
      case InspectionResult.CustomXmlMappingsResult result ->
          result.mappings().forEach(WorkbookInvariantChecks::requireCustomXmlMappingShape);
      case InspectionResult.CustomXmlExportResult result ->
          requireCustomXmlExportShape(result.export());
      case InspectionResult.NamedRangesResult result ->
          result.namedRanges().forEach(WorkbookInvariantChecks::requireNamedRangeShape);
      case InspectionResult.SheetSummaryResult result -> requireSheetSummaryShape(result.sheet());
      case InspectionResult.ArrayFormulasResult result ->
          result.arrayFormulas().forEach(WorkbookInvariantChecks::requireArrayFormulaShape);
      case InspectionResult.CellsResult result -> {
        require(result.sheetName() != null, "cells sheetName must not be null");
        require(!result.sheetName().isBlank(), "cells sheetName must not be blank");
        result.cells().forEach(WorkbookInvariantChecks::requireCellReportShape);
      }
      case InspectionResult.WindowResult result -> requireWindowShape(result.window());
      case InspectionResult.MergedRegionsResult result -> {
        require(result.sheetName() != null, "merged regions sheetName must not be null");
        require(!result.sheetName().isBlank(), "merged regions sheetName must not be blank");
        result
            .mergedRegions()
            .forEach(
                region ->
                    require(!region.range().isBlank(), "merged region range must not be blank"));
      }
      case InspectionResult.HyperlinksResult result -> {
        require(result.sheetName() != null, "hyperlinks sheetName must not be null");
        require(!result.sheetName().isBlank(), "hyperlinks sheetName must not be blank");
        result.hyperlinks().forEach(WorkbookInvariantChecks::requireHyperlinkEntryShape);
      }
      case InspectionResult.CommentsResult result -> {
        require(result.sheetName() != null, "comments sheetName must not be null");
        require(!result.sheetName().isBlank(), "comments sheetName must not be blank");
        result.comments().forEach(WorkbookInvariantChecks::requireCommentEntryShape);
      }
      case InspectionResult.DrawingObjectsResult result -> {
        require(result.sheetName() != null, "drawing objects sheetName must not be null");
        require(!result.sheetName().isBlank(), "drawing objects sheetName must not be blank");
        result.drawingObjects().forEach(WorkbookInvariantChecks::requireDrawingObjectShape);
      }
      case InspectionResult.ChartsResult result -> {
        require(result.sheetName() != null, "charts sheetName must not be null");
        require(!result.sheetName().isBlank(), "charts sheetName must not be blank");
        result.charts().forEach(WorkbookInvariantChecks::requireChartReportShape);
      }
      case InspectionResult.PivotTablesResult result ->
          result.pivotTables().forEach(WorkbookInvariantChecks::requirePivotTableShape);
      case InspectionResult.DrawingObjectPayloadResult result -> {
        require(result.sheetName() != null, "drawing payload sheetName must not be null");
        require(!result.sheetName().isBlank(), "drawing payload sheetName must not be blank");
        requireDrawingObjectPayloadShape(result.payload());
      }
      case InspectionResult.SheetLayoutResult result -> requireSheetLayoutShape(result.layout());
      case InspectionResult.PrintLayoutResult result -> requirePrintLayoutShape(result.layout());
      case InspectionResult.DataValidationsResult result -> {
        require(result.sheetName() != null, "data validations sheetName must not be null");
        require(!result.sheetName().isBlank(), "data validations sheetName must not be blank");
        result.validations().forEach(WorkbookInvariantChecks::requireDataValidationEntryShape);
      }
      case InspectionResult.ConditionalFormattingResult result -> {
        require(result.sheetName() != null, "conditional formatting sheetName must not be null");
        require(
            !result.sheetName().isBlank(), "conditional formatting sheetName must not be blank");
        result
            .conditionalFormattingBlocks()
            .forEach(WorkbookInvariantChecks::requireConditionalFormattingEntryShape);
      }
      case InspectionResult.AutofiltersResult result -> {
        require(result.sheetName() != null, "autofilters sheetName must not be null");
        require(!result.sheetName().isBlank(), "autofilters sheetName must not be blank");
        result.autofilters().forEach(WorkbookInvariantChecks::requireAutofilterEntryShape);
      }
      case InspectionResult.TablesResult result ->
          result.tables().forEach(WorkbookInvariantChecks::requireTableEntryShape);
      case InspectionResult.FormulaSurfaceResult result ->
          requireFormulaSurfaceShape(result.analysis());
      case InspectionResult.SheetSchemaResult result -> requireSheetSchemaShape(result.analysis());
      case InspectionResult.NamedRangeSurfaceResult result ->
          requireNamedRangeSurfaceShape(result.analysis());
      case InspectionResult.FormulaHealthResult result ->
          requireFormulaHealthShape(result.analysis());
      case InspectionResult.DataValidationHealthResult result ->
          requireDataValidationHealthShape(result.analysis());
      case InspectionResult.ConditionalFormattingHealthResult result ->
          requireConditionalFormattingHealthShape(result.analysis());
      case InspectionResult.AutofilterHealthResult result ->
          requireAutofilterHealthShape(result.analysis());
      case InspectionResult.TableHealthResult result -> requireTableHealthShape(result.analysis());
      case InspectionResult.PivotTableHealthResult result ->
          requirePivotTableHealthShape(result.analysis());
      case InspectionResult.HyperlinkHealthResult result ->
          requireHyperlinkHealthShape(result.analysis());
      case InspectionResult.NamedRangeHealthResult result ->
          requireNamedRangeHealthShape(result.analysis());
      case InspectionResult.WorkbookFindingsResult result ->
          requireWorkbookFindingsShape(result.analysis());
    }
  }

  private static void requireWorkbookSummaryShape(GridGrindResponse.WorkbookSummary workbook) {
    WorkbookInvariantWorkbookSurfaceChecks.requireWorkbookSummaryShape(workbook);
  }

  private static void requireSheetSummaryShape(GridGrindResponse.SheetSummaryReport sheet) {
    WorkbookInvariantWorkbookSurfaceChecks.requireSheetSummaryShape(sheet);
  }

  private static void requireDrawingObjectShape(DrawingObjectReport drawingObject) {
    WorkbookInvariantWorkbookSurfaceChecks.requireDrawingObjectShape(drawingObject);
  }

  private static void requireDrawingObjectPayloadShape(DrawingObjectPayloadReport payload) {
    WorkbookInvariantWorkbookSurfaceChecks.requireDrawingObjectPayloadShape(payload);
  }

  private static void requireChartReportShape(ChartReport chart) {
    WorkbookInvariantWorkbookSurfaceChecks.requireChartReportShape(chart);
  }

  private static void requireEngineWorkbookSummaryShape(
      dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary workbook) {
    WorkbookInvariantEngineShapeChecks.requireEngineWorkbookSummaryShape(workbook);
  }

  private static void requireEngineSheetSummaryShape(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary sheet) {
    WorkbookInvariantEngineShapeChecks.requireEngineSheetSummaryShape(sheet);
  }

  private static void requireEngineDrawingObjectShape(
      dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot drawingObject) {
    WorkbookInvariantEngineShapeChecks.requireEngineDrawingObjectShape(drawingObject);
  }

  private static void requireEngineChartShape(dev.erst.gridgrind.excel.ExcelChartSnapshot chart) {
    WorkbookInvariantEngineShapeChecks.requireEngineChartShape(chart);
  }

  private static void requireEnginePivotTableShape(
      dev.erst.gridgrind.excel.ExcelPivotTableSnapshot pivotTable) {
    WorkbookInvariantEngineShapeChecks.requireEnginePivotTableShape(pivotTable);
  }

  private static void requireWindowShape(GridGrindResponse.WindowReport window) {
    WorkbookInvariantAnalysisSurfaceChecks.requireWindowShape(window);
  }

  private static void requireHyperlinkEntryShape(GridGrindResponse.CellHyperlinkReport hyperlink) {
    WorkbookInvariantAnalysisSurfaceChecks.requireHyperlinkEntryShape(hyperlink);
  }

  private static void requireCommentEntryShape(GridGrindResponse.CellCommentReport comment) {
    WorkbookInvariantAnalysisSurfaceChecks.requireCommentEntryShape(comment);
  }

  private static void requireSheetLayoutShape(GridGrindResponse.SheetLayoutReport layout) {
    WorkbookInvariantAnalysisSurfaceChecks.requireSheetLayoutShape(layout);
  }

  private static void requirePrintLayoutShape(PrintLayoutReport layout) {
    WorkbookInvariantAnalysisSurfaceChecks.requirePrintLayoutShape(layout);
  }

  private static void requireDataValidationEntryShape(
      dev.erst.gridgrind.contract.dto.DataValidationEntryReport validation) {
    WorkbookInvariantAnalysisSurfaceChecks.requireDataValidationEntryShape(validation);
  }

  private static void requireAutofilterEntryShape(AutofilterEntryReport autofilter) {
    WorkbookInvariantAnalysisSurfaceChecks.requireAutofilterEntryShape(autofilter);
  }

  private static void requireConditionalFormattingEntryShape(
      ConditionalFormattingEntryReport conditionalFormatting) {
    WorkbookInvariantAnalysisSurfaceChecks.requireConditionalFormattingEntryShape(
        conditionalFormatting);
  }

  private static void requireTableEntryShape(TableEntryReport table) {
    WorkbookInvariantAnalysisSurfaceChecks.requireTableEntryShape(table);
  }

  private static void requirePivotTableShape(PivotTableReport pivotTable) {
    WorkbookInvariantAnalysisSurfaceChecks.requirePivotTableShape(pivotTable);
  }

  private static void requireFormulaSurfaceShape(GridGrindResponse.FormulaSurfaceReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireFormulaSurfaceShape(analysis);
  }

  private static void requireSheetSchemaShape(GridGrindResponse.SheetSchemaReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireSheetSchemaShape(analysis);
  }

  private static void requireNamedRangeSurfaceShape(
      GridGrindResponse.NamedRangeSurfaceReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireNamedRangeSurfaceShape(analysis);
  }

  private static void requireFormulaHealthShape(GridGrindResponse.FormulaHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireFormulaHealthShape(analysis);
  }

  private static void requireDataValidationHealthShape(
      dev.erst.gridgrind.contract.dto.DataValidationHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireDataValidationHealthShape(analysis);
  }

  private static void requireConditionalFormattingHealthShape(
      ConditionalFormattingHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireConditionalFormattingHealthShape(analysis);
  }

  private static void requireAutofilterHealthShape(AutofilterHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireAutofilterHealthShape(analysis);
  }

  private static void requireTableHealthShape(TableHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireTableHealthShape(analysis);
  }

  private static void requirePivotTableHealthShape(PivotTableHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requirePivotTableHealthShape(analysis);
  }

  private static void requireHyperlinkHealthShape(
      GridGrindResponse.HyperlinkHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireHyperlinkHealthShape(analysis);
  }

  private static void requireNamedRangeHealthShape(
      GridGrindResponse.NamedRangeHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireNamedRangeHealthShape(analysis);
  }

  private static void requireWorkbookFindingsShape(
      GridGrindResponse.WorkbookFindingsReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireWorkbookFindingsShape(analysis);
  }

  private static void requireCellReportShape(
      dev.erst.gridgrind.contract.dto.CellReport cellReport) {
    WorkbookInvariantCellSurfaceChecks.requireCellReportShape(cellReport);
  }

  private static void requireNamedRangeShape(GridGrindResponse.NamedRangeReport namedRange) {
    WorkbookInvariantCellSurfaceChecks.requireNamedRangeShape(namedRange);
  }

  private static void requireWorkbookProtectionShape(WorkbookProtectionReport protection) {
    WorkbookInvariantCellSurfaceChecks.requireWorkbookProtectionShape(protection);
  }

  private static void requirePackageSecurityShape(OoxmlPackageSecurityReport security) {
    WorkbookInvariantCellSurfaceChecks.requirePackageSecurityShape(security);
  }

  private static void requireCustomXmlMappingShape(CustomXmlMappingReport mapping) {
    WorkbookInvariantCellSurfaceChecks.requireCustomXmlMappingShape(mapping);
  }

  private static void requireCustomXmlExportShape(CustomXmlExportReport export) {
    WorkbookInvariantCellSurfaceChecks.requireCustomXmlExportShape(export);
  }

  private static void requireArrayFormulaShape(ArrayFormulaReport arrayFormula) {
    WorkbookInvariantCellSurfaceChecks.requireArrayFormulaShape(arrayFormula);
  }

  static void require(boolean condition, String message) {
    if (!condition) {
      throw new IllegalStateException(message);
    }
  }

  static void requireNonBlank(String value, String fieldName) {
    require(value != null, fieldName + " must not be null");
    require(!value.isBlank(), fieldName + " must not be blank");
  }

  static void requireBase64(String value, String fieldName) {
    require(value != null, fieldName + " must not be null");
    try {
      Base64.getDecoder().decode(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalStateException(fieldName + " must be valid base64", exception);
    }
  }

  static void requireNonBlank(TextSourceInput value, String fieldName) {
    require(value != null, fieldName + " must not be null");
    if (value instanceof TextSourceInput.Inline inline) {
      require(!inline.text().isBlank(), fieldName + " must not be blank");
    }
  }
}
