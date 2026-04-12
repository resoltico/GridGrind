package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.protocol.dto.AutofilterEntryReport;
import dev.erst.gridgrind.protocol.dto.AutofilterFilterColumnReport;
import dev.erst.gridgrind.protocol.dto.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.protocol.dto.AutofilterHealthReport;
import dev.erst.gridgrind.protocol.dto.AutofilterSortConditionReport;
import dev.erst.gridgrind.protocol.dto.AutofilterSortStateReport;
import dev.erst.gridgrind.protocol.dto.CellAlignmentReport;
import dev.erst.gridgrind.protocol.dto.CellBorderReport;
import dev.erst.gridgrind.protocol.dto.CellBorderSideReport;
import dev.erst.gridgrind.protocol.dto.CellColorReport;
import dev.erst.gridgrind.protocol.dto.CellFillReport;
import dev.erst.gridgrind.protocol.dto.CellFontReport;
import dev.erst.gridgrind.protocol.dto.CellGradientFillReport;
import dev.erst.gridgrind.protocol.dto.CellGradientStopReport;
import dev.erst.gridgrind.protocol.dto.CellProtectionReport;
import dev.erst.gridgrind.protocol.dto.ChartReport;
import dev.erst.gridgrind.protocol.dto.CommentAnchorReport;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.protocol.dto.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.protocol.dto.DifferentialBorderReport;
import dev.erst.gridgrind.protocol.dto.DifferentialBorderSideReport;
import dev.erst.gridgrind.protocol.dto.DifferentialStyleReport;
import dev.erst.gridgrind.protocol.dto.DrawingAnchorReport;
import dev.erst.gridgrind.protocol.dto.DrawingMarkerReport;
import dev.erst.gridgrind.protocol.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.protocol.dto.DrawingObjectReport;
import dev.erst.gridgrind.protocol.dto.FontHeightReport;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.dto.PaneReport;
import dev.erst.gridgrind.protocol.dto.PrintLayoutReport;
import dev.erst.gridgrind.protocol.dto.PrintMarginsReport;
import dev.erst.gridgrind.protocol.dto.PrintSetupReport;
import dev.erst.gridgrind.protocol.dto.RequestWarning;
import dev.erst.gridgrind.protocol.dto.RichTextRunReport;
import dev.erst.gridgrind.protocol.dto.TableColumnReport;
import dev.erst.gridgrind.protocol.dto.TableEntryReport;
import dev.erst.gridgrind.protocol.dto.TableHealthReport;
import dev.erst.gridgrind.protocol.dto.TableStyleReport;
import dev.erst.gridgrind.protocol.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;

/** Validates protocol responses and workbook state without depending on JUnit assertions. */
public final class WorkbookInvariantChecks {
  private WorkbookInvariantChecks() {}

  /** Requires the response shape to satisfy the protocol invariants the fuzzers rely on. */
  public static void requireResponseShape(GridGrindResponse response) {
    require(response != null, "response must not be null");
    require(response.protocolVersion() != null, "protocolVersion must not be null");

    switch (response) {
      case GridGrindResponse.Success success -> {
        require(success.persistence() != null, "persistence must not be null");
        requirePersistenceOutcomeShape(success.persistence());
        require(success.warnings() != null, "warnings must not be null");
        success.warnings().forEach(WorkbookInvariantChecks::requireRequestWarningShape);
        require(success.reads() != null, "reads must not be null");
        success.reads().forEach(WorkbookInvariantChecks::requireReadResultShape);
      }
      case GridGrindResponse.Failure failure -> {
        require(failure.problem() != null, "problem must not be null");
        require(failure.problem().code() != null, "problem code must not be null");
        require(failure.problem().category() != null, "problem category must not be null");
        require(failure.problem().recovery() != null, "problem recovery must not be null");
        require(failure.problem().title() != null, "problem title must not be null");
        require(failure.problem().message() != null, "problem message must not be null");
        require(failure.problem().resolution() != null, "problem resolution must not be null");
        require(failure.problem().context() != null, "problem context must not be null");
      }
    }
  }

  /**
   * Requires the response shape to agree with the request's source, reads, and persistence
   * contract.
   */
  public static void requireWorkflowOutcomeShape(
      GridGrindRequest request, GridGrindResponse response) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(response, "response must not be null");

    switch (response) {
      case GridGrindResponse.Failure _ -> {
        return;
      }
      case GridGrindResponse.Success success -> {
        requirePersistenceMatchesRequest(request.persistence(), success.persistence());
        require(
            success.reads().size() == request.reads().size(),
            "reads size must match the requested read count");
        for (int index = 0; index < request.reads().size(); index++) {
          requireReadMatchesRequest(request.reads().get(index), success.reads().get(index));
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
  }

  private static void requirePersistenceMatchesRequest(
      GridGrindRequest.WorkbookPersistence requestPersistence,
      GridGrindResponse.PersistenceOutcome persistenceOutcome) {
    switch (requestPersistence) {
      case GridGrindRequest.WorkbookPersistence.None _ -> {
        switch (persistenceOutcome) {
          case GridGrindResponse.PersistenceOutcome.NotSaved _ -> {}
          case GridGrindResponse.PersistenceOutcome.SavedAs _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
          case GridGrindResponse.PersistenceOutcome.Overwritten _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
        }
      }
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ -> {
        requirePersistenceOutcomeShape(persistenceOutcome);
        require(
            persistenceOutcome instanceof GridGrindResponse.PersistenceOutcome.Overwritten,
            "OVERWRITE persistence must return OVERWRITE outcome");
      }
      case GridGrindRequest.WorkbookPersistence.SaveAs _ -> {
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
    require(warning.operationIndex() >= 0, "warning operationIndex must not be negative");
    requireNonBlank(warning.operationType(), "warning operationType");
    requireNonBlank(warning.message(), "warning message");
  }

  private static void requireReadMatchesRequest(
      WorkbookReadOperation readOperation, WorkbookReadResult readResult) {
    require(
        readOperation.requestId().equals(readResult.requestId()),
        "read result requestId must match the request");
    require(
        SequenceIntrospection.readKind(readOperation).equals(readResultKind(readResult)),
        "read result kind must match the requested read kind");

    switch (readOperation) {
      case WorkbookReadOperation.GetWorkbookSummary _ -> {
        WorkbookReadResult.WorkbookSummaryResult result =
            (WorkbookReadResult.WorkbookSummaryResult) readResult;
        requireWorkbookSummaryShape(result.workbook());
      }
      case WorkbookReadOperation.GetWorkbookProtection _ ->
          requireWorkbookProtectionShape(
              ((WorkbookReadResult.WorkbookProtectionResult) readResult).protection());
      case WorkbookReadOperation.GetNamedRanges _ -> {
        WorkbookReadResult.NamedRangesResult result =
            (WorkbookReadResult.NamedRangesResult) readResult;
        result.namedRanges().forEach(WorkbookInvariantChecks::requireNamedRangeShape);
      }
      case WorkbookReadOperation.GetSheetSummary expected -> {
        WorkbookReadResult.SheetSummaryResult result =
            (WorkbookReadResult.SheetSummaryResult) readResult;
        requireSheetSummaryShape(result.sheet());
        require(
            expected.sheetName().equals(result.sheet().sheetName()),
            "sheet summary sheet mismatch");
      }
      case WorkbookReadOperation.GetCells expected -> {
        WorkbookReadResult.CellsResult result = (WorkbookReadResult.CellsResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "cells sheet mismatch");
        require(
            result.cells().size() == expected.addresses().size(),
            "cells result size must match requested addresses");
      }
      case WorkbookReadOperation.GetWindow expected -> {
        WorkbookReadResult.WindowResult result = (WorkbookReadResult.WindowResult) readResult;
        require(expected.sheetName().equals(result.window().sheetName()), "window sheet mismatch");
        require(
            expected.topLeftAddress().equals(result.window().topLeftAddress()),
            "window topLeftAddress mismatch");
        require(expected.rowCount() == result.window().rowCount(), "window rowCount mismatch");
        require(
            expected.columnCount() == result.window().columnCount(), "window columnCount mismatch");
      }
      case WorkbookReadOperation.GetMergedRegions expected -> {
        WorkbookReadResult.MergedRegionsResult result =
            (WorkbookReadResult.MergedRegionsResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "merged regions sheet mismatch");
      }
      case WorkbookReadOperation.GetHyperlinks expected -> {
        WorkbookReadResult.HyperlinksResult result =
            (WorkbookReadResult.HyperlinksResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "hyperlinks sheet mismatch");
      }
      case WorkbookReadOperation.GetComments expected -> {
        WorkbookReadResult.CommentsResult result = (WorkbookReadResult.CommentsResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "comments sheet mismatch");
      }
      case WorkbookReadOperation.GetDrawingObjects expected -> {
        WorkbookReadResult.DrawingObjectsResult result =
            (WorkbookReadResult.DrawingObjectsResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "drawing objects sheet mismatch");
        result.drawingObjects().forEach(WorkbookInvariantChecks::requireDrawingObjectShape);
      }
      case WorkbookReadOperation.GetCharts expected -> {
        WorkbookReadResult.ChartsResult result = (WorkbookReadResult.ChartsResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "charts sheet mismatch");
        result.charts().forEach(WorkbookInvariantChecks::requireChartReportShape);
      }
      case WorkbookReadOperation.GetDrawingObjectPayload expected -> {
        WorkbookReadResult.DrawingObjectPayloadResult result =
            (WorkbookReadResult.DrawingObjectPayloadResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "drawing payload sheet mismatch");
        requireDrawingObjectPayloadShape(result.payload());
        require(
            expected.objectName().equals(result.payload().name()),
            "drawing payload objectName mismatch");
      }
      case WorkbookReadOperation.GetSheetLayout expected -> {
        WorkbookReadResult.SheetLayoutResult result =
            (WorkbookReadResult.SheetLayoutResult) readResult;
        require(expected.sheetName().equals(result.layout().sheetName()), "layout sheet mismatch");
      }
      case WorkbookReadOperation.GetPrintLayout expected -> {
        WorkbookReadResult.PrintLayoutResult result =
            (WorkbookReadResult.PrintLayoutResult) readResult;
        require(
            expected.sheetName().equals(result.layout().sheetName()),
            "print layout sheet mismatch");
      }
      case WorkbookReadOperation.GetDataValidations expected -> {
        WorkbookReadResult.DataValidationsResult result =
            (WorkbookReadResult.DataValidationsResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "data validations sheet mismatch");
      }
      case WorkbookReadOperation.GetConditionalFormatting expected -> {
        WorkbookReadResult.ConditionalFormattingResult result =
            (WorkbookReadResult.ConditionalFormattingResult) readResult;
        require(
            expected.sheetName().equals(result.sheetName()),
            "conditional formatting sheet mismatch");
      }
      case WorkbookReadOperation.GetAutofilters expected -> {
        WorkbookReadResult.AutofiltersResult result =
            (WorkbookReadResult.AutofiltersResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "autofilters sheet mismatch");
      }
      case WorkbookReadOperation.GetTables _ -> {
        WorkbookReadResult.TablesResult result = (WorkbookReadResult.TablesResult) readResult;
        result.tables().forEach(WorkbookInvariantChecks::requireTableEntryShape);
      }
      case WorkbookReadOperation.GetFormulaSurface _ -> {
        WorkbookReadResult.FormulaSurfaceResult result =
            (WorkbookReadResult.FormulaSurfaceResult) readResult;
        require(result.analysis().sheets() != null, "formula surface sheets must not be null");
      }
      case WorkbookReadOperation.GetSheetSchema expected -> {
        WorkbookReadResult.SheetSchemaResult result =
            (WorkbookReadResult.SheetSchemaResult) readResult;
        require(
            expected.sheetName().equals(result.analysis().sheetName()), "schema sheet mismatch");
        require(
            expected.topLeftAddress().equals(result.analysis().topLeftAddress()),
            "schema topLeftAddress mismatch");
      }
      case WorkbookReadOperation.GetNamedRangeSurface _ -> {
        WorkbookReadResult.NamedRangeSurfaceResult result =
            (WorkbookReadResult.NamedRangeSurfaceResult) readResult;
        require(
            result.analysis().namedRanges() != null,
            "named range surface entries must not be null");
      }
      case WorkbookReadOperation.AnalyzeFormulaHealth _ ->
          requireFormulaHealthShape(
              ((WorkbookReadResult.FormulaHealthResult) readResult).analysis());
      case WorkbookReadOperation.AnalyzeDataValidationHealth _ ->
          requireDataValidationHealthShape(
              ((WorkbookReadResult.DataValidationHealthResult) readResult).analysis());
      case WorkbookReadOperation.AnalyzeConditionalFormattingHealth _ ->
          requireConditionalFormattingHealthShape(
              ((WorkbookReadResult.ConditionalFormattingHealthResult) readResult).analysis());
      case WorkbookReadOperation.AnalyzeAutofilterHealth _ ->
          requireAutofilterHealthShape(
              ((WorkbookReadResult.AutofilterHealthResult) readResult).analysis());
      case WorkbookReadOperation.AnalyzeTableHealth _ ->
          requireTableHealthShape(((WorkbookReadResult.TableHealthResult) readResult).analysis());
      case WorkbookReadOperation.AnalyzeHyperlinkHealth _ ->
          requireHyperlinkHealthShape(
              ((WorkbookReadResult.HyperlinkHealthResult) readResult).analysis());
      case WorkbookReadOperation.AnalyzeNamedRangeHealth _ ->
          requireNamedRangeHealthShape(
              ((WorkbookReadResult.NamedRangeHealthResult) readResult).analysis());
      case WorkbookReadOperation.AnalyzeWorkbookFindings _ ->
          requireWorkbookFindingsShape(
              ((WorkbookReadResult.WorkbookFindingsResult) readResult).analysis());
    }
  }

  private static String readResultKind(WorkbookReadResult readResult) {
    return switch (readResult) {
      case WorkbookReadResult.WorkbookSummaryResult _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadResult.WorkbookProtectionResult _ -> "GET_WORKBOOK_PROTECTION";
      case WorkbookReadResult.NamedRangesResult _ -> "GET_NAMED_RANGES";
      case WorkbookReadResult.SheetSummaryResult _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadResult.CellsResult _ -> "GET_CELLS";
      case WorkbookReadResult.WindowResult _ -> "GET_WINDOW";
      case WorkbookReadResult.MergedRegionsResult _ -> "GET_MERGED_REGIONS";
      case WorkbookReadResult.HyperlinksResult _ -> "GET_HYPERLINKS";
      case WorkbookReadResult.CommentsResult _ -> "GET_COMMENTS";
      case WorkbookReadResult.DrawingObjectsResult _ -> "GET_DRAWING_OBJECTS";
      case WorkbookReadResult.ChartsResult _ -> "GET_CHARTS";
      case WorkbookReadResult.DrawingObjectPayloadResult _ -> "GET_DRAWING_OBJECT_PAYLOAD";
      case WorkbookReadResult.SheetLayoutResult _ -> "GET_SHEET_LAYOUT";
      case WorkbookReadResult.PrintLayoutResult _ -> "GET_PRINT_LAYOUT";
      case WorkbookReadResult.DataValidationsResult _ -> "GET_DATA_VALIDATIONS";
      case WorkbookReadResult.ConditionalFormattingResult _ -> "GET_CONDITIONAL_FORMATTING";
      case WorkbookReadResult.AutofiltersResult _ -> "GET_AUTOFILTERS";
      case WorkbookReadResult.TablesResult _ -> "GET_TABLES";
      case WorkbookReadResult.FormulaSurfaceResult _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadResult.SheetSchemaResult _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadResult.NamedRangeSurfaceResult _ -> "GET_NAMED_RANGE_SURFACE";
      case WorkbookReadResult.FormulaHealthResult _ -> "ANALYZE_FORMULA_HEALTH";
      case WorkbookReadResult.DataValidationHealthResult _ -> "ANALYZE_DATA_VALIDATION_HEALTH";
      case WorkbookReadResult.ConditionalFormattingHealthResult _ ->
          "ANALYZE_CONDITIONAL_FORMATTING_HEALTH";
      case WorkbookReadResult.AutofilterHealthResult _ -> "ANALYZE_AUTOFILTER_HEALTH";
      case WorkbookReadResult.TableHealthResult _ -> "ANALYZE_TABLE_HEALTH";
      case WorkbookReadResult.HyperlinkHealthResult _ -> "ANALYZE_HYPERLINK_HEALTH";
      case WorkbookReadResult.NamedRangeHealthResult _ -> "ANALYZE_NAMED_RANGE_HEALTH";
      case WorkbookReadResult.WorkbookFindingsResult _ -> "ANALYZE_WORKBOOK_FINDINGS";
    };
  }

  private static void requireReadResultShape(WorkbookReadResult readResult) {
    require(readResult.requestId() != null, "read requestId must not be null");
    require(!readResult.requestId().isBlank(), "read requestId must not be blank");

    switch (readResult) {
      case WorkbookReadResult.WorkbookSummaryResult result ->
          requireWorkbookSummaryShape(result.workbook());
      case WorkbookReadResult.WorkbookProtectionResult result ->
          requireWorkbookProtectionShape(result.protection());
      case WorkbookReadResult.NamedRangesResult result ->
          result.namedRanges().forEach(WorkbookInvariantChecks::requireNamedRangeShape);
      case WorkbookReadResult.SheetSummaryResult result -> requireSheetSummaryShape(result.sheet());
      case WorkbookReadResult.CellsResult result -> {
        require(result.sheetName() != null, "cells sheetName must not be null");
        require(!result.sheetName().isBlank(), "cells sheetName must not be blank");
        result.cells().forEach(WorkbookInvariantChecks::requireCellReportShape);
      }
      case WorkbookReadResult.WindowResult result -> requireWindowShape(result.window());
      case WorkbookReadResult.MergedRegionsResult result -> {
        require(result.sheetName() != null, "merged regions sheetName must not be null");
        require(!result.sheetName().isBlank(), "merged regions sheetName must not be blank");
        result
            .mergedRegions()
            .forEach(
                region ->
                    require(!region.range().isBlank(), "merged region range must not be blank"));
      }
      case WorkbookReadResult.HyperlinksResult result -> {
        require(result.sheetName() != null, "hyperlinks sheetName must not be null");
        require(!result.sheetName().isBlank(), "hyperlinks sheetName must not be blank");
        result.hyperlinks().forEach(WorkbookInvariantChecks::requireHyperlinkEntryShape);
      }
      case WorkbookReadResult.CommentsResult result -> {
        require(result.sheetName() != null, "comments sheetName must not be null");
        require(!result.sheetName().isBlank(), "comments sheetName must not be blank");
        result.comments().forEach(WorkbookInvariantChecks::requireCommentEntryShape);
      }
      case WorkbookReadResult.DrawingObjectsResult result -> {
        require(result.sheetName() != null, "drawing objects sheetName must not be null");
        require(!result.sheetName().isBlank(), "drawing objects sheetName must not be blank");
        result.drawingObjects().forEach(WorkbookInvariantChecks::requireDrawingObjectShape);
      }
      case WorkbookReadResult.ChartsResult result -> {
        require(result.sheetName() != null, "charts sheetName must not be null");
        require(!result.sheetName().isBlank(), "charts sheetName must not be blank");
        result.charts().forEach(WorkbookInvariantChecks::requireChartReportShape);
      }
      case WorkbookReadResult.DrawingObjectPayloadResult result -> {
        require(result.sheetName() != null, "drawing payload sheetName must not be null");
        require(!result.sheetName().isBlank(), "drawing payload sheetName must not be blank");
        requireDrawingObjectPayloadShape(result.payload());
      }
      case WorkbookReadResult.SheetLayoutResult result -> requireSheetLayoutShape(result.layout());
      case WorkbookReadResult.PrintLayoutResult result -> requirePrintLayoutShape(result.layout());
      case WorkbookReadResult.DataValidationsResult result -> {
        require(result.sheetName() != null, "data validations sheetName must not be null");
        require(!result.sheetName().isBlank(), "data validations sheetName must not be blank");
        result.validations().forEach(WorkbookInvariantChecks::requireDataValidationEntryShape);
      }
      case WorkbookReadResult.ConditionalFormattingResult result -> {
        require(result.sheetName() != null, "conditional formatting sheetName must not be null");
        require(
            !result.sheetName().isBlank(), "conditional formatting sheetName must not be blank");
        result
            .conditionalFormattingBlocks()
            .forEach(WorkbookInvariantChecks::requireConditionalFormattingEntryShape);
      }
      case WorkbookReadResult.AutofiltersResult result -> {
        require(result.sheetName() != null, "autofilters sheetName must not be null");
        require(!result.sheetName().isBlank(), "autofilters sheetName must not be blank");
        result.autofilters().forEach(WorkbookInvariantChecks::requireAutofilterEntryShape);
      }
      case WorkbookReadResult.TablesResult result ->
          result.tables().forEach(WorkbookInvariantChecks::requireTableEntryShape);
      case WorkbookReadResult.FormulaSurfaceResult result ->
          requireFormulaSurfaceShape(result.analysis());
      case WorkbookReadResult.SheetSchemaResult result ->
          requireSheetSchemaShape(result.analysis());
      case WorkbookReadResult.NamedRangeSurfaceResult result ->
          requireNamedRangeSurfaceShape(result.analysis());
      case WorkbookReadResult.FormulaHealthResult result ->
          requireFormulaHealthShape(result.analysis());
      case WorkbookReadResult.DataValidationHealthResult result ->
          requireDataValidationHealthShape(result.analysis());
      case WorkbookReadResult.ConditionalFormattingHealthResult result ->
          requireConditionalFormattingHealthShape(result.analysis());
      case WorkbookReadResult.AutofilterHealthResult result ->
          requireAutofilterHealthShape(result.analysis());
      case WorkbookReadResult.TableHealthResult result ->
          requireTableHealthShape(result.analysis());
      case WorkbookReadResult.HyperlinkHealthResult result ->
          requireHyperlinkHealthShape(result.analysis());
      case WorkbookReadResult.NamedRangeHealthResult result ->
          requireNamedRangeHealthShape(result.analysis());
      case WorkbookReadResult.WorkbookFindingsResult result ->
          requireWorkbookFindingsShape(result.analysis());
    }
  }

  private static void requireWorkbookSummaryShape(GridGrindResponse.WorkbookSummary workbook) {
    require(workbook != null, "workbook summary must not be null");
    require(workbook.sheetCount() >= 0, "sheetCount must not be negative");
    require(workbook.namedRangeCount() >= 0, "namedRangeCount must not be negative");
    require(workbook.sheetNames() != null, "sheetNames must not be null");
    require(
        workbook.sheetCount() == workbook.sheetNames().size(),
        "sheetCount must match sheetNames size");
    workbook
        .sheetNames()
        .forEach(
            sheetName ->
                require(sheetName != null && !sheetName.isBlank(), "sheetName must not be blank"));
    switch (workbook) {
      case GridGrindResponse.WorkbookSummary.Empty empty -> {
        require(empty.sheetCount() == 0, "empty workbook summary must have sheetCount 0");
        require(empty.sheetNames().isEmpty(), "empty workbook summary must have no sheet names");
      }
      case GridGrindResponse.WorkbookSummary.WithSheets withSheets -> {
        require(
            withSheets.sheetCount() > 0,
            "non-empty workbook summary must have positive sheetCount");
        requireNonBlank(withSheets.activeSheetName(), "activeSheetName");
        require(
            withSheets.sheetNames().contains(withSheets.activeSheetName()),
            "activeSheetName must be present in sheetNames");
        require(withSheets.selectedSheetNames() != null, "selectedSheetNames must not be null");
        require(!withSheets.selectedSheetNames().isEmpty(), "selectedSheetNames must not be empty");
        require(
            withSheets.selectedSheetNames().size()
                == new HashSet<>(withSheets.selectedSheetNames()).size(),
            "selectedSheetNames must be unique");
        withSheets
            .selectedSheetNames()
            .forEach(
                selectedSheetName -> {
                  requireNonBlank(selectedSheetName, "selectedSheetName");
                  require(
                      withSheets.sheetNames().contains(selectedSheetName),
                      "selectedSheetNames must be present in sheetNames");
                });
      }
    }
  }

  private static void requireSheetSummaryShape(GridGrindResponse.SheetSummaryReport sheet) {
    require(sheet.sheetName() != null, "sheetName must not be null");
    require(!sheet.sheetName().isBlank(), "sheetName must not be blank");
    require(sheet.visibility() != null, "visibility must not be null");
    require(sheet.protection() != null, "protection must not be null");
    require(sheet.physicalRowCount() >= 0, "physicalRowCount must not be negative");
    require(sheet.lastRowIndex() >= -1, "lastRowIndex must be greater than or equal to -1");
    require(sheet.lastColumnIndex() >= -1, "lastColumnIndex must be greater than or equal to -1");
    switch (sheet.protection()) {
      case GridGrindResponse.SheetProtectionReport.Unprotected _ -> {}
      case GridGrindResponse.SheetProtectionReport.Protected protectedReport ->
          require(protectedReport.settings() != null, "protected sheet settings must not be null");
    }
  }

  private static void requireDrawingObjectShape(DrawingObjectReport drawingObject) {
    require(drawingObject != null, "drawing object must not be null");
    requireNonBlank(drawingObject.name(), "drawing object name");
    requireDrawingAnchorShape(drawingObject.anchor());
    switch (drawingObject) {
      case DrawingObjectReport.Picture picture -> requirePictureDrawingObjectShape(picture);
      case DrawingObjectReport.Chart chart -> requireChartDrawingObjectShape(chart);
      case DrawingObjectReport.Shape shape -> requireShapeDrawingObjectShape(shape);
      case DrawingObjectReport.EmbeddedObject embeddedObject ->
          requireEmbeddedDrawingObjectShape(embeddedObject);
    }
  }

  private static void requirePictureDrawingObjectShape(DrawingObjectReport.Picture picture) {
    requireNonBlank(picture.contentType(), "picture contentType");
    requireNonBlank(picture.sha256(), "picture sha256");
    require(picture.byteSize() > 0L, "picture byteSize must be positive");
    if (picture.widthPixels() != null) {
      require(picture.widthPixels() >= 0, "picture widthPixels must not be negative");
    }
    if (picture.heightPixels() != null) {
      require(picture.heightPixels() >= 0, "picture heightPixels must not be negative");
    }
    if (picture.description() != null) {
      require(!picture.description().isBlank(), "picture description must not be blank");
    }
  }

  private static void requireChartDrawingObjectShape(DrawingObjectReport.Chart chart) {
    require(chart.plotTypeTokens() != null, "chart plotTypeTokens must not be null");
    chart.plotTypeTokens().forEach(token -> requireNonBlank(token, "chart plotTypeToken"));
    require(chart.title() != null, "chart title must not be null");
  }

  private static void requireShapeDrawingObjectShape(DrawingObjectReport.Shape shape) {
    if (shape.presetGeometryToken() != null) {
      require(
          !shape.presetGeometryToken().isBlank(), "shape presetGeometryToken must not be blank");
    }
    if (shape.text() != null) {
      require(!shape.text().isBlank(), "shape text must not be blank");
    }
    require(shape.childCount() >= 0, "shape childCount must not be negative");
  }

  private static void requireEmbeddedDrawingObjectShape(
      DrawingObjectReport.EmbeddedObject embeddedObject) {
    requireNonBlank(embeddedObject.contentType(), "embedded object contentType");
    requireNonBlank(embeddedObject.sha256(), "embedded object sha256");
    require(embeddedObject.byteSize() > 0L, "embedded object byteSize must be positive");
    if (embeddedObject.label() != null) {
      require(!embeddedObject.label().isBlank(), "embedded object label must not be blank");
    }
    if (embeddedObject.fileName() != null) {
      require(!embeddedObject.fileName().isBlank(), "embedded object fileName must not be blank");
    }
    if (embeddedObject.command() != null) {
      require(!embeddedObject.command().isBlank(), "embedded object command must not be blank");
    }
    if (embeddedObject.previewByteSize() != null) {
      require(
          embeddedObject.previewByteSize() > 0L,
          "embedded object previewByteSize must be positive");
      require(
          embeddedObject.previewFormat() != null,
          "embedded object previewByteSize requires previewFormat");
    }
    if (embeddedObject.previewSha256() != null) {
      require(
          !embeddedObject.previewSha256().isBlank(),
          "embedded object previewSha256 must not be blank");
      require(
          embeddedObject.previewFormat() != null,
          "embedded object previewSha256 requires previewFormat");
    }
  }

  private static void requireDrawingObjectPayloadShape(DrawingObjectPayloadReport payload) {
    require(payload != null, "drawing payload must not be null");
    requireNonBlank(payload.name(), "drawing payload name");
    requireNonBlank(payload.contentType(), "drawing payload contentType");
    requireNonBlank(payload.sha256(), "drawing payload sha256");
    requireNonBlank(payload.base64Data(), "drawing payload base64Data");
    switch (payload) {
      case DrawingObjectPayloadReport.Picture picture -> {
        requireNonBlank(picture.fileName(), "picture payload fileName");
        if (picture.description() != null) {
          require(
              !picture.description().isBlank(), "picture payload description must not be blank");
        }
      }
      case DrawingObjectPayloadReport.EmbeddedObject embeddedObject -> {
        if (embeddedObject.fileName() != null) {
          require(
              !embeddedObject.fileName().isBlank(), "embedded payload fileName must not be blank");
        }
        if (embeddedObject.label() != null) {
          require(!embeddedObject.label().isBlank(), "embedded payload label must not be blank");
        }
        if (embeddedObject.command() != null) {
          require(
              !embeddedObject.command().isBlank(), "embedded payload command must not be blank");
        }
      }
    }
  }

  private static void requireDrawingAnchorShape(DrawingAnchorReport anchor) {
    require(anchor != null, "drawing anchor must not be null");
    switch (anchor) {
      case DrawingAnchorReport.TwoCell twoCell -> {
        requireDrawingMarkerShape(twoCell.from());
        requireDrawingMarkerShape(twoCell.to());
      }
      case DrawingAnchorReport.OneCell oneCell -> {
        requireDrawingMarkerShape(oneCell.from());
        require(oneCell.widthEmu() > 0L, "one-cell widthEmu must be positive");
        require(oneCell.heightEmu() > 0L, "one-cell heightEmu must be positive");
      }
      case DrawingAnchorReport.Absolute absolute -> {
        require(absolute.xEmu() >= 0L, "absolute xEmu must not be negative");
        require(absolute.yEmu() >= 0L, "absolute yEmu must not be negative");
        require(absolute.widthEmu() > 0L, "absolute widthEmu must be positive");
        require(absolute.heightEmu() > 0L, "absolute heightEmu must be positive");
      }
    }
  }

  private static void requireDrawingMarkerShape(DrawingMarkerReport marker) {
    require(marker != null, "drawing marker must not be null");
    require(marker.columnIndex() >= 0, "drawing marker columnIndex must not be negative");
    require(marker.rowIndex() >= 0, "drawing marker rowIndex must not be negative");
    require(marker.dx() >= 0, "drawing marker dx must not be negative");
    require(marker.dy() >= 0, "drawing marker dy must not be negative");
  }

  private static void requireChartReportShape(ChartReport chart) {
    require(chart != null, "chart report must not be null");
    requireNonBlank(chart.name(), "chart name");
    requireDrawingAnchorShape(chart.anchor());
    switch (chart) {
      case ChartReport.Bar bar -> {
        require(bar.title() != null, "bar chart title must not be null");
        require(bar.legend() != null, "bar chart legend must not be null");
        require(bar.displayBlanksAs() != null, "bar chart displayBlanksAs must not be null");
        require(bar.barDirection() != null, "bar chart barDirection must not be null");
        bar.axes().forEach(WorkbookInvariantChecks::requireChartAxisShape);
        bar.series().forEach(WorkbookInvariantChecks::requireChartSeriesShape);
      }
      case ChartReport.Line line -> {
        require(line.title() != null, "line chart title must not be null");
        require(line.legend() != null, "line chart legend must not be null");
        require(line.displayBlanksAs() != null, "line chart displayBlanksAs must not be null");
        line.axes().forEach(WorkbookInvariantChecks::requireChartAxisShape);
        line.series().forEach(WorkbookInvariantChecks::requireChartSeriesShape);
      }
      case ChartReport.Pie pie -> {
        require(pie.title() != null, "pie chart title must not be null");
        require(pie.legend() != null, "pie chart legend must not be null");
        require(pie.displayBlanksAs() != null, "pie chart displayBlanksAs must not be null");
        if (pie.firstSliceAngle() != null) {
          require(
              pie.firstSliceAngle() >= 0 && pie.firstSliceAngle() <= 360,
              "pie chart firstSliceAngle must be between 0 and 360");
        }
        pie.series().forEach(WorkbookInvariantChecks::requireChartSeriesShape);
      }
      case ChartReport.Unsupported unsupported -> {
        unsupported
            .plotTypeTokens()
            .forEach(token -> requireNonBlank(token, "unsupported chart plotTypeToken"));
        requireNonBlank(unsupported.detail(), "unsupported chart detail");
      }
    }
  }

  private static void requireChartAxisShape(ChartReport.Axis axis) {
    require(axis != null, "chart axis must not be null");
    require(axis.kind() != null, "chart axis kind must not be null");
    require(axis.position() != null, "chart axis position must not be null");
    require(axis.crosses() != null, "chart axis crosses must not be null");
  }

  private static void requireChartSeriesShape(ChartReport.Series series) {
    require(series != null, "chart series must not be null");
    requireChartTitleShape(series.title());
    requireChartDataSourceShape(series.categories());
    requireChartDataSourceShape(series.values());
  }

  private static void requireChartTitleShape(ChartReport.Title title) {
    require(title != null, "chart title must not be null");
    switch (title) {
      case ChartReport.Title.None _ -> {}
      case ChartReport.Title.Text text ->
          require(text.text() != null, "chart title text must not be null");
      case ChartReport.Title.Formula formula -> {
        requireNonBlank(formula.formula(), "chart title formula");
        require(formula.cachedText() != null, "chart title cachedText must not be null");
      }
    }
  }

  private static void requireChartDataSourceShape(ChartReport.DataSource source) {
    require(source != null, "chart data source must not be null");
    switch (source) {
      case ChartReport.DataSource.StringReference reference -> {
        requireNonBlank(reference.formula(), "chart string-reference formula");
        require(
            reference.cachedValues() != null,
            "chart string-reference cachedValues must not be null");
      }
      case ChartReport.DataSource.NumericReference reference -> {
        requireNonBlank(reference.formula(), "chart numeric-reference formula");
        require(
            reference.cachedValues() != null,
            "chart numeric-reference cachedValues must not be null");
        if (reference.formatCode() != null) {
          require(
              !reference.formatCode().isBlank(),
              "chart numeric-reference formatCode must not be blank");
        }
      }
      case ChartReport.DataSource.StringLiteral literal ->
          require(literal.values() != null, "chart string-literal values must not be null");
      case ChartReport.DataSource.NumericLiteral literal -> {
        require(literal.values() != null, "chart numeric-literal values must not be null");
        if (literal.formatCode() != null) {
          require(
              !literal.formatCode().isBlank(),
              "chart numeric-literal formatCode must not be blank");
        }
      }
    }
  }

  private static void requireEngineWorkbookSummaryShape(
      dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary workbook) {
    require(workbook != null, "engine workbook summary must not be null");
    require(workbook.sheetCount() >= 0, "engine sheetCount must not be negative");
    require(workbook.namedRangeCount() >= 0, "engine namedRangeCount must not be negative");
    require(workbook.sheetNames() != null, "engine sheetNames must not be null");
    require(
        workbook.sheetCount() == workbook.sheetNames().size(),
        "engine sheetCount must match sheetNames size");
    require(
        workbook.sheetNames().size() == new HashSet<>(workbook.sheetNames()).size(),
        "engine sheet names must be unique");
    workbook.sheetNames().forEach(sheetName -> requireNonBlank(sheetName, "engine sheetName"));
    switch (workbook) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty empty -> {
        require(empty.sheetCount() == 0, "engine empty workbook summary must have sheetCount 0");
        require(
            empty.sheetNames().isEmpty(), "engine empty workbook summary must have no sheet names");
      }
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets withSheets -> {
        require(
            withSheets.sheetCount() > 0,
            "engine non-empty workbook summary must have positive sheetCount");
        requireNonBlank(withSheets.activeSheetName(), "engine activeSheetName");
        require(
            withSheets.sheetNames().contains(withSheets.activeSheetName()),
            "engine activeSheetName must be present in sheetNames");
        require(
            !withSheets.selectedSheetNames().isEmpty(),
            "engine selectedSheetNames must not be empty");
        require(
            withSheets.selectedSheetNames().size()
                == new HashSet<>(withSheets.selectedSheetNames()).size(),
            "engine selectedSheetNames must be unique");
        withSheets
            .selectedSheetNames()
            .forEach(
                selectedSheetName -> {
                  requireNonBlank(selectedSheetName, "engine selectedSheetName");
                  require(
                      withSheets.sheetNames().contains(selectedSheetName),
                      "engine selectedSheetNames must be present in sheetNames");
                });
      }
    }
  }

  private static void requireEngineSheetSummaryShape(
      dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary sheet) {
    requireNonBlank(sheet.sheetName(), "engine sheetName");
    require(sheet.visibility() != null, "engine visibility must not be null");
    require(sheet.protection() != null, "engine protection must not be null");
    require(sheet.physicalRowCount() >= 0, "engine physicalRowCount must not be negative");
    require(sheet.lastRowIndex() >= -1, "engine lastRowIndex must be greater than or equal to -1");
    require(
        sheet.lastColumnIndex() >= -1,
        "engine lastColumnIndex must be greater than or equal to -1");
    switch (sheet.protection()) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Unprotected _ -> {}
      case dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Protected protectedSheet ->
          require(
              protectedSheet.settings() != null,
              "engine protected sheet settings must not be null");
    }
  }

  private static void requireEngineDrawingObjectShape(
      dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot drawingObject) {
    require(drawingObject != null, "engine drawing object must not be null");
    requireNonBlank(drawingObject.name(), "engine drawing object name");
    requireEngineDrawingAnchorShape(drawingObject.anchor());
    switch (drawingObject) {
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.Picture picture -> {
        requireNonBlank(picture.contentType(), "engine picture contentType");
        requireNonBlank(picture.sha256(), "engine picture sha256");
        require(picture.byteSize() > 0L, "engine picture byteSize must be positive");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.Chart chart -> {
        require(chart.plotTypeTokens() != null, "engine chart plotTypeTokens must not be null");
        chart
            .plotTypeTokens()
            .forEach(token -> requireNonBlank(token, "engine chart plotTypeToken"));
        require(chart.title() != null, "engine chart title must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.Shape shape -> {
        if (shape.presetGeometryToken() != null) {
          require(
              !shape.presetGeometryToken().isBlank(),
              "engine shape presetGeometryToken must not be blank");
        }
        if (shape.text() != null) {
          require(!shape.text().isBlank(), "engine shape text must not be blank");
        }
        require(shape.childCount() >= 0, "engine shape childCount must not be negative");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.EmbeddedObject embeddedObject -> {
        requireNonBlank(embeddedObject.contentType(), "engine embedded contentType");
        requireNonBlank(embeddedObject.sha256(), "engine embedded sha256");
        require(embeddedObject.byteSize() > 0L, "engine embedded byteSize must be positive");
      }
    }
  }

  private static void requireEngineChartShape(dev.erst.gridgrind.excel.ExcelChartSnapshot chart) {
    require(chart != null, "engine chart must not be null");
    requireNonBlank(chart.name(), "engine chart name");
    requireEngineDrawingAnchorShape(chart.anchor());
    switch (chart) {
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Bar bar -> {
        require(bar.title() != null, "engine bar chart title must not be null");
        require(bar.legend() != null, "engine bar chart legend must not be null");
        require(bar.displayBlanksAs() != null, "engine bar chart displayBlanksAs must not be null");
        require(bar.barDirection() != null, "engine bar chart barDirection must not be null");
        bar.axes().forEach(WorkbookInvariantChecks::requireEngineChartAxisShape);
        bar.series().forEach(WorkbookInvariantChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Line line -> {
        require(line.title() != null, "engine line chart title must not be null");
        require(line.legend() != null, "engine line chart legend must not be null");
        require(
            line.displayBlanksAs() != null, "engine line chart displayBlanksAs must not be null");
        line.axes().forEach(WorkbookInvariantChecks::requireEngineChartAxisShape);
        line.series().forEach(WorkbookInvariantChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Pie pie -> {
        require(pie.title() != null, "engine pie chart title must not be null");
        require(pie.legend() != null, "engine pie chart legend must not be null");
        require(pie.displayBlanksAs() != null, "engine pie chart displayBlanksAs must not be null");
        if (pie.firstSliceAngle() != null) {
          require(
              pie.firstSliceAngle() >= 0 && pie.firstSliceAngle() <= 360,
              "engine pie chart firstSliceAngle must be between 0 and 360");
        }
        pie.series().forEach(WorkbookInvariantChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Unsupported unsupported -> {
        unsupported
            .plotTypeTokens()
            .forEach(token -> requireNonBlank(token, "engine unsupported chart plotTypeToken"));
        requireNonBlank(unsupported.detail(), "engine unsupported chart detail");
      }
    }
  }

  private static void requireEngineChartAxisShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.Axis axis) {
    require(axis != null, "engine chart axis must not be null");
    require(axis.kind() != null, "engine chart axis kind must not be null");
    require(axis.position() != null, "engine chart axis position must not be null");
    require(axis.crosses() != null, "engine chart axis crosses must not be null");
  }

  private static void requireEngineChartSeriesShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.Series series) {
    require(series != null, "engine chart series must not be null");
    requireEngineChartTitleShape(series.title());
    requireEngineChartDataSourceShape(series.categories());
    requireEngineChartDataSourceShape(series.values());
  }

  private static void requireEngineChartTitleShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.Title title) {
    require(title != null, "engine chart title must not be null");
    switch (title) {
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Title.None _ -> {}
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Title.Text text ->
          require(text.text() != null, "engine chart title text must not be null");
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Title.Formula formula -> {
        requireNonBlank(formula.formula(), "engine chart title formula");
        require(formula.cachedText() != null, "engine chart title cachedText must not be null");
      }
    }
  }

  private static void requireEngineChartDataSourceShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource source) {
    require(source != null, "engine chart data source must not be null");
    switch (source) {
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.StringReference reference -> {
        requireNonBlank(reference.formula(), "engine chart string-reference formula");
        require(
            reference.cachedValues() != null,
            "engine chart string-reference cachedValues must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.NumericReference reference -> {
        requireNonBlank(reference.formula(), "engine chart numeric-reference formula");
        require(
            reference.cachedValues() != null,
            "engine chart numeric-reference cachedValues must not be null");
        if (reference.formatCode() != null) {
          require(
              !reference.formatCode().isBlank(),
              "engine chart numeric-reference formatCode must not be blank");
        }
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.StringLiteral literal ->
          require(literal.values() != null, "engine chart string-literal values must not be null");
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.NumericLiteral literal -> {
        require(literal.values() != null, "engine chart numeric-literal values must not be null");
        if (literal.formatCode() != null) {
          require(
              !literal.formatCode().isBlank(),
              "engine chart numeric-literal formatCode must not be blank");
        }
      }
    }
  }

  private static void requireEngineDrawingAnchorShape(
      dev.erst.gridgrind.excel.ExcelDrawingAnchor anchor) {
    require(anchor != null, "engine drawing anchor must not be null");
    switch (anchor) {
      case dev.erst.gridgrind.excel.ExcelDrawingAnchor.TwoCell twoCell -> {
        requireEngineDrawingMarkerShape(twoCell.from());
        requireEngineDrawingMarkerShape(twoCell.to());
        require(twoCell.behavior() != null, "engine two-cell anchor behavior must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingAnchor.OneCell oneCell -> {
        requireEngineDrawingMarkerShape(oneCell.from());
        require(oneCell.widthEmu() > 0L, "engine one-cell widthEmu must be positive");
        require(oneCell.heightEmu() > 0L, "engine one-cell heightEmu must be positive");
        require(oneCell.behavior() != null, "engine one-cell anchor behavior must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingAnchor.Absolute absolute -> {
        require(absolute.xEmu() >= 0L, "engine absolute xEmu must not be negative");
        require(absolute.yEmu() >= 0L, "engine absolute yEmu must not be negative");
        require(absolute.widthEmu() > 0L, "engine absolute widthEmu must be positive");
        require(absolute.heightEmu() > 0L, "engine absolute heightEmu must be positive");
        require(absolute.behavior() != null, "engine absolute anchor behavior must not be null");
      }
    }
  }

  private static void requireEngineDrawingMarkerShape(
      dev.erst.gridgrind.excel.ExcelDrawingMarker marker) {
    require(marker != null, "engine drawing marker must not be null");
    require(marker.columnIndex() >= 0, "engine drawing marker columnIndex must not be negative");
    require(marker.rowIndex() >= 0, "engine drawing marker rowIndex must not be negative");
    require(marker.dx() >= 0, "engine drawing marker dx must not be negative");
    require(marker.dy() >= 0, "engine drawing marker dy must not be negative");
  }

  private static void requireWindowShape(GridGrindResponse.WindowReport window) {
    require(window.sheetName() != null, "window sheetName must not be null");
    require(!window.sheetName().isBlank(), "window sheetName must not be blank");
    require(window.topLeftAddress() != null, "window topLeftAddress must not be null");
    require(!window.topLeftAddress().isBlank(), "window topLeftAddress must not be blank");
    require(window.rows() != null, "window rows must not be null");
    require(window.rows().size() == window.rowCount(), "window rows size must match rowCount");
    window
        .rows()
        .forEach(
            row -> {
              require(row.rowIndex() >= 0, "window row index must not be negative");
              require(row.cells() != null, "window row cells must not be null");
              require(
                  row.cells().size() == window.columnCount(),
                  "window row cells size must match columnCount");
              row.cells().forEach(WorkbookInvariantChecks::requireCellReportShape);
            });
  }

  private static void requireHyperlinkEntryShape(GridGrindResponse.CellHyperlinkReport hyperlink) {
    require(hyperlink.address() != null, "hyperlink address must not be null");
    require(!hyperlink.address().isBlank(), "hyperlink address must not be blank");
    require(hyperlink.hyperlink() != null, "hyperlink metadata must not be null");
    requireHyperlinkShape(hyperlink.hyperlink());
  }

  private static void requireCommentEntryShape(GridGrindResponse.CellCommentReport comment) {
    require(comment.address() != null, "comment address must not be null");
    require(!comment.address().isBlank(), "comment address must not be blank");
    require(comment.comment() != null, "comment metadata must not be null");
    requireCommentReportShape(comment.comment());
  }

  private static void requireSheetLayoutShape(GridGrindResponse.SheetLayoutReport layout) {
    require(layout.sheetName() != null, "layout sheetName must not be null");
    require(!layout.sheetName().isBlank(), "layout sheetName must not be blank");
    require(layout.pane() != null, "pane must not be null");
    require(
        layout.zoomPercent() >= 10 && layout.zoomPercent() <= 400,
        "zoomPercent must be between 10 and 400 inclusive");
    switch (layout.pane()) {
      case PaneReport.None _ -> {}
      case PaneReport.Frozen frozen -> {
        require(frozen.splitColumn() >= 0, "splitColumn must not be negative");
        require(frozen.splitRow() >= 0, "splitRow must not be negative");
        require(frozen.leftmostColumn() >= 0, "leftmostColumn must not be negative");
        require(frozen.topRow() >= 0, "topRow must not be negative");
      }
      case PaneReport.Split split -> {
        require(split.xSplitPosition() >= 0, "xSplitPosition must not be negative");
        require(split.ySplitPosition() >= 0, "ySplitPosition must not be negative");
        require(split.leftmostColumn() >= 0, "leftmostColumn must not be negative");
        require(split.topRow() >= 0, "topRow must not be negative");
        require(split.activePane() != null, "activePane must not be null");
      }
    }
    layout
        .columns()
        .forEach(
            column -> {
              require(column.columnIndex() >= 0, "columnIndex must not be negative");
              require(
                  Double.isFinite(column.widthCharacters()) && column.widthCharacters() > 0.0d,
                  "column width must be finite and greater than 0");
            });
    layout
        .rows()
        .forEach(
            row -> {
              require(row.rowIndex() >= 0, "rowIndex must not be negative");
              require(
                  Double.isFinite(row.heightPoints()) && row.heightPoints() > 0.0d,
                  "row height must be finite and greater than 0");
            });
  }

  private static void requirePrintLayoutShape(PrintLayoutReport layout) {
    require(layout.sheetName() != null, "print layout sheetName must not be null");
    require(!layout.sheetName().isBlank(), "print layout sheetName must not be blank");
    require(layout.printArea() != null, "printArea must not be null");
    require(layout.orientation() != null, "orientation must not be null");
    require(layout.scaling() != null, "scaling must not be null");
    require(layout.repeatingRows() != null, "repeatingRows must not be null");
    require(layout.repeatingColumns() != null, "repeatingColumns must not be null");
    require(layout.header() != null, "header must not be null");
    require(layout.footer() != null, "footer must not be null");
    requirePrintSetupShape(layout.setup());
  }

  private static void requireDataValidationEntryShape(
      dev.erst.gridgrind.protocol.dto.DataValidationEntryReport validation) {
    require(validation.ranges() != null, "data validation ranges must not be null");
    require(!validation.ranges().isEmpty(), "data validation ranges must not be empty");
    validation.ranges().forEach(range -> requireNonBlank(range, "data validation range"));

    switch (validation) {
      case dev.erst.gridgrind.protocol.dto.DataValidationEntryReport.Supported supported ->
          requireSupportedDataValidationShape(supported.validation());
      case dev.erst.gridgrind.protocol.dto.DataValidationEntryReport.Unsupported unsupported -> {
        requireNonBlank(unsupported.kind(), "data validation kind");
        requireNonBlank(unsupported.detail(), "data validation detail");
      }
    }
  }

  private static void requireAutofilterEntryShape(AutofilterEntryReport autofilter) {
    requireNonBlank(autofilter.range(), "autofilter range");
    require(autofilter.filterColumns() != null, "autofilter filterColumns must not be null");
    autofilter.filterColumns().forEach(WorkbookInvariantChecks::requireAutofilterFilterColumnShape);
    if (autofilter.sortState() != null) {
      requireAutofilterSortStateShape(autofilter.sortState());
    }
    switch (autofilter) {
      case AutofilterEntryReport.SheetOwned _ -> {}
      case AutofilterEntryReport.TableOwned tableOwned ->
          requireNonBlank(tableOwned.tableName(), "autofilter table name");
    }
  }

  private static void requireConditionalFormattingEntryShape(
      ConditionalFormattingEntryReport conditionalFormatting) {
    require(
        conditionalFormatting.ranges() != null, "conditional formatting ranges must not be null");
    require(
        !conditionalFormatting.ranges().isEmpty(),
        "conditional formatting ranges must not be empty");
    conditionalFormatting
        .ranges()
        .forEach(range -> requireNonBlank(range, "conditional formatting range"));
    require(conditionalFormatting.rules() != null, "conditional formatting rules must not be null");
    require(
        !conditionalFormatting.rules().isEmpty(), "conditional formatting rules must not be empty");
    conditionalFormatting
        .rules()
        .forEach(WorkbookInvariantChecks::requireConditionalFormattingRuleShape);
  }

  private static void requireTableEntryShape(TableEntryReport table) {
    requireNonBlank(table.name(), "table name");
    requireNonBlank(table.sheetName(), "table sheetName");
    requireNonBlank(table.range(), "table range");
    require(table.headerRowCount() >= 0, "table headerRowCount must not be negative");
    require(table.totalsRowCount() >= 0, "table totalsRowCount must not be negative");
    require(table.columnNames() != null, "table columnNames must not be null");
    require(table.columns() != null, "table columns must not be null");
    require(
        table.columnNames().size() == table.columns().size(),
        "table columnNames size must match columns size");
    table
        .columnNames()
        .forEach(columnName -> require(columnName != null, "table column name must not be null"));
    for (int index = 0; index < table.columns().size(); index++) {
      requireTableColumnShape(table.columns().get(index));
      require(
          table.columnNames().get(index).equals(table.columns().get(index).name()),
          "table columnNames must align with columns");
    }
    require(table.style() != null, "table style must not be null");
    requireTableStyleShape(table.style());
  }

  private static void requireConditionalFormattingRuleShape(ConditionalFormattingRuleReport rule) {
    require(rule.priority() > 0, "conditional formatting priority must be greater than 0");
    switch (rule) {
      case ConditionalFormattingRuleReport.FormulaRule formulaRule -> {
        requireNonBlank(formulaRule.formula(), "conditional formatting formula");
        require(formulaRule.style() != null, "conditional formatting style must not be null");
        requireDifferentialStyleShape(formulaRule.style());
      }
      case ConditionalFormattingRuleReport.CellValueRule cellValueRule -> {
        require(
            cellValueRule.operator() != null, "conditional formatting operator must not be null");
        requireNonBlank(cellValueRule.formula1(), "conditional formatting formula1");
        if (cellValueRule.formula2() != null) {
          requireNonBlank(cellValueRule.formula2(), "conditional formatting formula2");
        }
        if (cellValueRule.style() != null) {
          requireDifferentialStyleShape(cellValueRule.style());
        }
      }
      case ConditionalFormattingRuleReport.ColorScaleRule colorScaleRule -> {
        require(
            colorScaleRule.thresholds() != null,
            "conditional formatting thresholds must not be null");
        require(
            !colorScaleRule.thresholds().isEmpty(),
            "conditional formatting thresholds must not be empty");
        colorScaleRule
            .thresholds()
            .forEach(WorkbookInvariantChecks::requireConditionalFormattingThresholdShape);
        require(colorScaleRule.colors() != null, "conditional formatting colors must not be null");
        require(
            !colorScaleRule.colors().isEmpty(), "conditional formatting colors must not be empty");
        colorScaleRule
            .colors()
            .forEach(color -> requireNonBlank(color, "conditional formatting color"));
      }
      case ConditionalFormattingRuleReport.DataBarRule dataBarRule -> {
        requireNonBlank(dataBarRule.color(), "conditional formatting color");
        requireConditionalFormattingThresholdShape(dataBarRule.minThreshold());
        requireConditionalFormattingThresholdShape(dataBarRule.maxThreshold());
        require(
            dataBarRule.widthMin() >= 0, "conditional formatting widthMin must not be negative");
        require(
            dataBarRule.widthMax() >= 0, "conditional formatting widthMax must not be negative");
      }
      case ConditionalFormattingRuleReport.IconSetRule iconSetRule -> {
        require(iconSetRule.iconSet() != null, "conditional formatting iconSet must not be null");
        require(
            iconSetRule.thresholds() != null, "conditional formatting thresholds must not be null");
        require(
            !iconSetRule.thresholds().isEmpty(),
            "conditional formatting thresholds must not be empty");
        iconSetRule
            .thresholds()
            .forEach(WorkbookInvariantChecks::requireConditionalFormattingThresholdShape);
      }
      case ConditionalFormattingRuleReport.Top10Rule top10Rule -> {
        require(top10Rule.rank() >= 0, "conditional formatting rank must not be negative");
        if (top10Rule.style() != null) {
          requireDifferentialStyleShape(top10Rule.style());
        }
      }
      case ConditionalFormattingRuleReport.UnsupportedRule unsupportedRule -> {
        requireNonBlank(unsupportedRule.kind(), "conditional formatting kind");
        requireNonBlank(unsupportedRule.detail(), "conditional formatting detail");
      }
    }
  }

  private static void requireConditionalFormattingThresholdShape(
      ConditionalFormattingThresholdReport threshold) {
    require(threshold != null, "conditional formatting threshold must not be null");
    require(threshold.type() != null, "conditional formatting threshold type must not be null");
  }

  private static void requireDifferentialStyleShape(DifferentialStyleReport style) {
    require(style != null, "conditional formatting style must not be null");
    if (style.numberFormat() != null) {
      requireNonBlank(style.numberFormat(), "conditional formatting numberFormat");
    }
    if (style.fontHeight() != null) {
      requireFontHeightShape(style.fontHeight());
    }
    if (style.fontColor() != null) {
      requireNonBlank(style.fontColor(), "conditional formatting fontColor");
    }
    if (style.fillColor() != null) {
      requireNonBlank(style.fillColor(), "conditional formatting fillColor");
    }
    if (style.border() != null) {
      requireDifferentialBorderShape(style.border());
    }
    require(
        style.unsupportedFeatures() != null,
        "conditional formatting unsupportedFeatures must not be null");
    style
        .unsupportedFeatures()
        .forEach(
            feature ->
                require(
                    feature != null,
                    "conditional formatting unsupported feature must not be null"));
  }

  private static void requireDifferentialBorderShape(DifferentialBorderReport border) {
    require(border != null, "conditional formatting border must not be null");
    if (border.all() != null) {
      requireDifferentialBorderSideShape(border.all());
    }
    if (border.top() != null) {
      requireDifferentialBorderSideShape(border.top());
    }
    if (border.right() != null) {
      requireDifferentialBorderSideShape(border.right());
    }
    if (border.bottom() != null) {
      requireDifferentialBorderSideShape(border.bottom());
    }
    if (border.left() != null) {
      requireDifferentialBorderSideShape(border.left());
    }
  }

  private static void requireDifferentialBorderSideShape(DifferentialBorderSideReport side) {
    require(side != null, "conditional formatting border side must not be null");
    require(side.style() != null, "conditional formatting border style must not be null");
    if (side.color() != null) {
      requireNonBlank(side.color(), "conditional formatting border color");
    }
  }

  private static void requireTableStyleShape(TableStyleReport style) {
    switch (style) {
      case TableStyleReport.None _ -> {}
      case TableStyleReport.Named named -> requireNonBlank(named.name(), "table style name");
    }
  }

  private static void requireSupportedDataValidationShape(
      dev.erst.gridgrind.protocol.dto.DataValidationEntryReport.DataValidationDefinitionReport
          validation) {
    require(validation != null, "data validation definition must not be null");
    require(validation.rule() != null, "data validation rule must not be null");
    switch (validation.rule()) {
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.ExplicitList explicitList -> {
        require(explicitList.values() != null, "explicit list values must not be null");
        explicitList.values().forEach(value -> requireNonBlank(value, "explicit list value"));
      }
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.FormulaList formulaList ->
          requireNonBlank(formulaList.formula(), "formula list formula");
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.WholeNumber wholeNumber ->
          requireComparisonRuleShape(wholeNumber.operator(), wholeNumber.formula1());
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.DecimalNumber decimalNumber ->
          requireComparisonRuleShape(decimalNumber.operator(), decimalNumber.formula1());
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.DateRule dateRule ->
          requireComparisonRuleShape(dateRule.operator(), dateRule.formula1());
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.TimeRule timeRule ->
          requireComparisonRuleShape(timeRule.operator(), timeRule.formula1());
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.TextLength textLength ->
          requireComparisonRuleShape(textLength.operator(), textLength.formula1());
      case dev.erst.gridgrind.protocol.dto.DataValidationRuleInput.CustomFormula customFormula ->
          requireNonBlank(customFormula.formula(), "custom validation formula");
    }
    if (validation.prompt() != null) {
      requireNonBlank(validation.prompt().title(), "data validation prompt title");
      requireNonBlank(validation.prompt().text(), "data validation prompt text");
    }
    if (validation.errorAlert() != null) {
      require(
          validation.errorAlert().style() != null, "data validation error style must not be null");
      requireNonBlank(validation.errorAlert().title(), "data validation error title");
      requireNonBlank(validation.errorAlert().text(), "data validation error text");
    }
  }

  private static void requireComparisonRuleShape(Object operator, String formula1) {
    require(operator != null, "comparison operator must not be null");
    requireNonBlank(formula1, "comparison formula1");
  }

  private static void requireFormulaSurfaceShape(GridGrindResponse.FormulaSurfaceReport analysis) {
    require(analysis.totalFormulaCellCount() >= 0, "totalFormulaCellCount must not be negative");
    analysis
        .sheets()
        .forEach(
            sheet -> {
              require(sheet.sheetName() != null, "formula surface sheetName must not be null");
              require(!sheet.sheetName().isBlank(), "formula surface sheetName must not be blank");
              require(sheet.formulaCellCount() >= 0, "formulaCellCount must not be negative");
              require(
                  sheet.distinctFormulaCount() >= 0, "distinctFormulaCount must not be negative");
              sheet
                  .formulas()
                  .forEach(
                      formula -> {
                        require(formula.formula() != null, "formula pattern must not be null");
                        require(!formula.formula().isBlank(), "formula pattern must not be blank");
                        require(
                            formula.occurrenceCount() > 0,
                            "occurrenceCount must be greater than 0");
                        require(formula.addresses() != null, "formula addresses must not be null");
                      });
            });
  }

  private static void requireSheetSchemaShape(GridGrindResponse.SheetSchemaReport analysis) {
    require(analysis.sheetName() != null, "schema sheetName must not be null");
    require(!analysis.sheetName().isBlank(), "schema sheetName must not be blank");
    require(analysis.topLeftAddress() != null, "schema topLeftAddress must not be null");
    require(!analysis.topLeftAddress().isBlank(), "schema topLeftAddress must not be blank");
    require(analysis.rowCount() > 0, "schema rowCount must be greater than 0");
    require(analysis.columnCount() > 0, "schema columnCount must be greater than 0");
    require(analysis.dataRowCount() >= 0, "schema dataRowCount must not be negative");
    analysis
        .columns()
        .forEach(
            column -> {
              require(column.columnIndex() >= 0, "schema columnIndex must not be negative");
              require(column.columnAddress() != null, "schema columnAddress must not be null");
              require(!column.columnAddress().isBlank(), "schema columnAddress must not be blank");
              require(
                  column.headerDisplayValue() != null,
                  "schema headerDisplayValue must not be null");
              require(
                  column.populatedCellCount() >= 0,
                  "schema populatedCellCount must not be negative");
              require(column.blankCellCount() >= 0, "schema blankCellCount must not be negative");
              column
                  .observedTypes()
                  .forEach(
                      typeCount -> {
                        require(typeCount.type() != null, "type count type must not be null");
                        require(!typeCount.type().isBlank(), "type count type must not be blank");
                        require(typeCount.count() > 0, "type count must be greater than 0");
                      });
            });
  }

  private static void requireNamedRangeSurfaceShape(
      GridGrindResponse.NamedRangeSurfaceReport analysis) {
    require(analysis.workbookScopedCount() >= 0, "workbookScopedCount must not be negative");
    require(analysis.sheetScopedCount() >= 0, "sheetScopedCount must not be negative");
    require(analysis.rangeBackedCount() >= 0, "rangeBackedCount must not be negative");
    require(analysis.formulaBackedCount() >= 0, "formulaBackedCount must not be negative");
    analysis
        .namedRanges()
        .forEach(
            namedRange -> {
              require(namedRange.name() != null, "named range name must not be null");
              require(!namedRange.name().isBlank(), "named range name must not be blank");
              require(namedRange.scope() != null, "named range scope must not be null");
              require(
                  namedRange.refersToFormula() != null,
                  "named range refersToFormula must not be null");
              require(namedRange.kind() != null, "named range kind must not be null");
            });
  }

  private static void requireFormulaHealthShape(GridGrindResponse.FormulaHealthReport analysis) {
    require(
        analysis.checkedFormulaCellCount() >= 0, "checkedFormulaCellCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireDataValidationHealthShape(
      dev.erst.gridgrind.protocol.dto.DataValidationHealthReport analysis) {
    require(analysis.checkedValidationCount() >= 0, "checkedValidationCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireConditionalFormattingHealthShape(
      ConditionalFormattingHealthReport analysis) {
    require(
        analysis.checkedConditionalFormattingBlockCount() >= 0,
        "checkedConditionalFormattingBlockCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireAutofilterHealthShape(AutofilterHealthReport analysis) {
    require(analysis.checkedAutofilterCount() >= 0, "checkedAutofilterCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireTableHealthShape(TableHealthReport analysis) {
    require(analysis.checkedTableCount() >= 0, "checkedTableCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireHyperlinkHealthShape(
      GridGrindResponse.HyperlinkHealthReport analysis) {
    require(analysis.checkedHyperlinkCount() >= 0, "checkedHyperlinkCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireNamedRangeHealthShape(
      GridGrindResponse.NamedRangeHealthReport analysis) {
    require(analysis.checkedNamedRangeCount() >= 0, "checkedNamedRangeCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireWorkbookFindingsShape(
      GridGrindResponse.WorkbookFindingsReport analysis) {
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireAnalysisSummaryShape(
      GridGrindResponse.AnalysisSummaryReport summary,
      List<GridGrindResponse.AnalysisFindingReport> findings) {
    require(summary != null, "analysis summary must not be null");
    require(findings != null, "analysis findings must not be null");
    require(summary.totalCount() >= 0, "analysis totalCount must not be negative");
    require(summary.errorCount() >= 0, "analysis errorCount must not be negative");
    require(summary.warningCount() >= 0, "analysis warningCount must not be negative");
    require(summary.infoCount() >= 0, "analysis infoCount must not be negative");
    require(
        summary.totalCount() == findings.size(), "analysis totalCount must match findings size");
    require(
        summary.totalCount() == summary.errorCount() + summary.warningCount() + summary.infoCount(),
        "analysis totalCount must equal error + warning + info");
    findings.forEach(WorkbookInvariantChecks::requireAnalysisFindingShape);
  }

  private static void requireAnalysisFindingShape(GridGrindResponse.AnalysisFindingReport finding) {
    require(finding.code() != null, "analysis finding code must not be null");
    require(finding.severity() != null, "analysis finding severity must not be null");
    requireNonBlank(finding.title(), "analysis title");
    requireNonBlank(finding.message(), "analysis message");
    require(finding.location() != null, "analysis location must not be null");
    require(finding.evidence() != null, "analysis evidence must not be null");
    finding.evidence().forEach(evidence -> requireNonBlank(evidence, "analysis evidence"));

    switch (finding.location()) {
      case GridGrindResponse.AnalysisLocationReport.Workbook _ -> {}
      case GridGrindResponse.AnalysisLocationReport.Sheet sheet ->
          requireNonBlank(sheet.sheetName(), "analysis sheetName");
      case GridGrindResponse.AnalysisLocationReport.Cell cell -> {
        requireNonBlank(cell.sheetName(), "analysis sheetName");
        requireNonBlank(cell.address(), "analysis address");
      }
      case GridGrindResponse.AnalysisLocationReport.Range range -> {
        requireNonBlank(range.sheetName(), "analysis sheetName");
        requireNonBlank(range.range(), "analysis range");
      }
      case GridGrindResponse.AnalysisLocationReport.NamedRange namedRange -> {
        requireNonBlank(namedRange.name(), "analysis named range");
        require(namedRange.scope() != null, "analysis named range scope must not be null");
      }
    }
  }

  private static void requireCellReportShape(GridGrindResponse.CellReport cellReport) {
    require(cellReport.address() != null, "cell address must not be null");
    require(!cellReport.address().isBlank(), "cell address must not be blank");
    require(cellReport.declaredType() != null, "declaredType must not be null");
    require(cellReport.effectiveType() != null, "effectiveType must not be null");
    require(cellReport.displayValue() != null, "displayValue must not be null");
    requireCellStyleShape(cellReport.style());

    switch (cellReport) {
      case GridGrindResponse.CellReport.BlankReport _ -> {}
      case GridGrindResponse.CellReport.TextReport text -> {
        require(text.stringValue() != null, "stringValue must not be null");
        if (text.richText() != null) {
          require(!text.richText().isEmpty(), "richText must not be empty");
          StringBuilder builder = new StringBuilder();
          for (var run : text.richText()) {
            require(run.text() != null, "richText run text must not be null");
            require(!run.text().isEmpty(), "richText run text must not be empty");
            requireCellFontShape(run.font());
            builder.append(run.text());
          }
          require(
              text.stringValue().equals(builder.toString()),
              "richText run text must concatenate to stringValue");
        }
      }
      case GridGrindResponse.CellReport.NumberReport number ->
          require(number.numberValue() != null, "numberValue must not be null");
      case GridGrindResponse.CellReport.BooleanReport bool ->
          require(bool.booleanValue() != null, "booleanValue must not be null");
      case GridGrindResponse.CellReport.ErrorReport error ->
          require(error.errorValue() != null, "errorValue must not be null");
      case GridGrindResponse.CellReport.FormulaReport formula -> {
        require(formula.formula() != null, "formula must not be null");
        requireCellReportShape(formula.evaluation());
      }
    }
    if (cellReport.hyperlink() != null) {
      requireHyperlinkShape(cellReport.hyperlink());
    }
    if (cellReport.comment() != null) {
      requireCommentReportShape(cellReport.comment());
    }
  }

  private static void requireCommentReportShape(GridGrindResponse.CommentReport comment) {
    require(comment.text() != null, "comment text must not be null");
    require(comment.author() != null, "comment author must not be null");
    require(!comment.text().isBlank(), "comment text must not be blank");
    require(!comment.author().isBlank(), "comment author must not be blank");
    if (comment.runs() != null) {
      require(!comment.runs().isEmpty(), "comment runs must not be empty");
      StringBuilder builder = new StringBuilder();
      for (RichTextRunReport run : comment.runs()) {
        require(run != null, "comment runs must not contain null values");
        require(run.text() != null, "comment run text must not be null");
        require(!run.text().isEmpty(), "comment run text must not be empty");
        requireCellFontShape(run.font());
        builder.append(run.text());
      }
      require(builder.toString().equals(comment.text()), "comment runs must concatenate to text");
    }
    if (comment.anchor() != null) {
      requireCommentAnchorShape(comment.anchor());
    }
  }

  private static void requireNamedRangeShape(GridGrindResponse.NamedRangeReport namedRange) {
    require(namedRange.name() != null, "namedRange name must not be null");
    require(!namedRange.name().isBlank(), "namedRange name must not be blank");
    require(namedRange.scope() != null, "namedRange scope must not be null");
    require(namedRange.refersToFormula() != null, "namedRange formula must not be null");

    switch (namedRange) {
      case GridGrindResponse.NamedRangeReport.RangeReport range -> {
        require(range.target() != null, "namedRange target must not be null");
        require(range.target().sheetName() != null, "namedRange target sheet must not be null");
        require(range.target().range() != null, "namedRange target range must not be null");
        require(!range.target().sheetName().isBlank(), "namedRange target sheet must not be blank");
        require(!range.target().range().isBlank(), "namedRange target range must not be blank");
      }
      case GridGrindResponse.NamedRangeReport.FormulaReport _ -> {}
    }
  }

  private static void requireHyperlinkShape(HyperlinkTarget hyperlink) {
    require(hyperlink != null, "hyperlink must not be null");
    switch (hyperlink) {
      case HyperlinkTarget.Url url -> {
        require(url.target() != null, "hyperlink target must not be null");
        require(!url.target().isBlank(), "hyperlink target must not be blank");
        require(
            !url.target().regionMatches(true, 0, "file:", 0, 5),
            "URL hyperlink targets must not use file: schemes");
        require(
            !url.target().regionMatches(true, 0, "mailto:", 0, 7),
            "URL hyperlink targets must not use mailto: schemes");
      }
      case HyperlinkTarget.Email email -> {
        require(email.email() != null, "hyperlink email must not be null");
        require(!email.email().isBlank(), "hyperlink email must not be blank");
        require(
            !email.email().regionMatches(true, 0, "mailto:", 0, 7),
            "EMAIL hyperlink targets must omit the mailto: prefix");
      }
      case HyperlinkTarget.File file -> {
        require(file.path() != null, "hyperlink path must not be null");
        require(!file.path().isBlank(), "hyperlink path must not be blank");
        require(
            !file.path().regionMatches(true, 0, "file:", 0, 5),
            "FILE hyperlink targets must be normalized path strings");
      }
      case HyperlinkTarget.Document document -> {
        require(document.target() != null, "hyperlink target must not be null");
        require(!document.target().isBlank(), "hyperlink target must not be blank");
      }
    }
  }

  private static void requireCellStyleShape(GridGrindResponse.CellStyleReport style) {
    require(style != null, "style must not be null");
    require(style.numberFormat() != null, "numberFormat must not be null");
    requireCellAlignmentShape(style.alignment());
    requireCellFontShape(style.font());
    requireCellFillShape(style.fill());
    requireCellBorderShape(style.border());
    requireCellProtectionShape(style.protection());
  }

  private static void requireCellAlignmentShape(CellAlignmentReport alignment) {
    require(alignment != null, "alignment must not be null");
    require(alignment.horizontalAlignment() != null, "horizontalAlignment must not be null");
    require(alignment.verticalAlignment() != null, "verticalAlignment must not be null");
    require(
        alignment.textRotation() >= 0 && alignment.textRotation() <= 180,
        "textRotation must be between 0 and 180 inclusive");
    require(
        alignment.indentation() >= 0 && alignment.indentation() <= 250,
        "indentation must be between 0 and 250 inclusive");
  }

  private static void requireCellFontShape(CellFontReport font) {
    require(font != null, "font must not be null");
    require(font.fontName() != null, "fontName must not be null");
    require(!font.fontName().isBlank(), "fontName must not be blank");
    requireFontHeightShape(font.fontHeight());
    if (font.fontColor() != null) {
      requireCellColorShape(font.fontColor(), "fontColor");
    }
  }

  private static void requireCellFillShape(CellFillReport fill) {
    require(fill != null, "fill must not be null");
    require(fill.pattern() != null, "fill pattern must not be null");
    if (fill.foregroundColor() != null) {
      requireCellColorShape(fill.foregroundColor(), "fill foregroundColor");
    }
    if (fill.backgroundColor() != null) {
      requireCellColorShape(fill.backgroundColor(), "fill backgroundColor");
    }
    if (fill.gradient() != null) {
      requireCellGradientFillShape(fill.gradient());
      require(
          fill.foregroundColor() == null && fill.backgroundColor() == null,
          "gradient fills must not carry flat colors");
    }
    if (fill.pattern() == ExcelFillPattern.NONE && fill.gradient() == null) {
      require(
          fill.foregroundColor() == null && fill.backgroundColor() == null,
          "fill pattern NONE must not carry colors");
    }
    if (fill.pattern() == ExcelFillPattern.SOLID && fill.gradient() == null) {
      require(fill.backgroundColor() == null, "SOLID fills must not carry backgroundColor");
    }
  }

  private static void requireCellBorderShape(CellBorderReport border) {
    require(border != null, "border must not be null");
    requireCellBorderSideShape(border.top(), "top");
    requireCellBorderSideShape(border.right(), "right");
    requireCellBorderSideShape(border.bottom(), "bottom");
    requireCellBorderSideShape(border.left(), "left");
  }

  private static void requireCellBorderSideShape(CellBorderSideReport side, String label) {
    require(side != null, label + " border side must not be null");
    require(side.style() != null, label + " border style must not be null");
    if (side.color() != null) {
      requireCellColorShape(side.color(), label + " border color");
    }
  }

  private static void requireWorkbookProtectionShape(WorkbookProtectionReport protection) {
    require(protection != null, "workbook protection must not be null");
  }

  private static void requireCommentAnchorShape(CommentAnchorReport anchor) {
    require(anchor.firstColumn() >= 0, "comment anchor firstColumn must not be negative");
    require(anchor.firstRow() >= 0, "comment anchor firstRow must not be negative");
    require(anchor.lastColumn() >= anchor.firstColumn(), "comment anchor columns must be ordered");
    require(anchor.lastRow() >= anchor.firstRow(), "comment anchor rows must be ordered");
  }

  private static void requirePrintSetupShape(PrintSetupReport setup) {
    require(setup != null, "print setup must not be null");
    requirePrintMarginsShape(setup.margins());
    require(setup.paperSize() >= 0, "print setup paperSize must not be negative");
    require(setup.copies() >= 0, "print setup copies must not be negative");
    require(setup.firstPageNumber() >= 0, "print setup firstPageNumber must not be negative");
    require(setup.rowBreaks() != null, "print setup rowBreaks must not be null");
    require(setup.columnBreaks() != null, "print setup columnBreaks must not be null");
    setup
        .rowBreaks()
        .forEach(rowBreak -> require(rowBreak >= 0, "print setup rowBreak must not be negative"));
    setup
        .columnBreaks()
        .forEach(
            columnBreak ->
                require(columnBreak >= 0, "print setup columnBreak must not be negative"));
  }

  private static void requirePrintMarginsShape(PrintMarginsReport margins) {
    require(margins != null, "print margins must not be null");
  }

  private static void requireAutofilterFilterColumnShape(
      AutofilterFilterColumnReport filterColumn) {
    require(filterColumn != null, "autofilter filterColumn must not be null");
    require(filterColumn.columnId() >= 0L, "autofilter columnId must not be negative");
    requireAutofilterCriterionShape(filterColumn.criterion());
  }

  private static void requireAutofilterCriterionShape(AutofilterFilterCriterionReport criterion) {
    require(criterion != null, "autofilter criterion must not be null");
    switch (criterion) {
      case AutofilterFilterCriterionReport.Values values -> {
        require(values.values() != null, "autofilter values must not be null");
        values
            .values()
            .forEach(value -> require(value != null, "autofilter value must not be null"));
      }
      case AutofilterFilterCriterionReport.Custom custom -> {
        require(custom.conditions() != null, "autofilter custom conditions must not be null");
        require(!custom.conditions().isEmpty(), "autofilter custom conditions must not be empty");
        custom
            .conditions()
            .forEach(
                condition -> {
                  require(condition != null, "autofilter custom condition must not be null");
                  requireNonBlank(condition.operator(), "autofilter custom operator");
                  requireNonBlank(condition.value(), "autofilter custom value");
                });
      }
      case AutofilterFilterCriterionReport.Dynamic dynamic -> {
        requireNonBlank(dynamic.type(), "autofilter dynamic type");
        if (dynamic.value() != null) {
          require(Double.isFinite(dynamic.value()), "autofilter dynamic value must be finite");
        }
        if (dynamic.maxValue() != null) {
          require(
              Double.isFinite(dynamic.maxValue()), "autofilter dynamic maxValue must be finite");
        }
      }
      case AutofilterFilterCriterionReport.Top10 top10 -> {
        require(Double.isFinite(top10.value()), "autofilter top10 value must be finite");
        require(top10.value() >= 0.0d, "autofilter top10 value must not be negative");
        if (top10.filterValue() != null) {
          require(
              Double.isFinite(top10.filterValue()), "autofilter top10 filterValue must be finite");
        }
      }
      case AutofilterFilterCriterionReport.Color color -> {
        if (color.color() != null) {
          requireCellColorShape(color.color(), "autofilter color");
        }
      }
      case AutofilterFilterCriterionReport.Icon icon -> {
        requireNonBlank(icon.iconSet(), "autofilter iconSet");
        require(icon.iconId() >= 0, "autofilter iconId must not be negative");
      }
    }
  }

  private static void requireAutofilterSortStateShape(AutofilterSortStateReport sortState) {
    require(sortState != null, "autofilter sortState must not be null");
    requireNonBlank(sortState.range(), "autofilter sortState range");
    require(sortState.conditions() != null, "autofilter sortState conditions must not be null");
    sortState.conditions().forEach(WorkbookInvariantChecks::requireAutofilterSortConditionShape);
  }

  private static void requireAutofilterSortConditionShape(AutofilterSortConditionReport condition) {
    require(condition != null, "autofilter sort condition must not be null");
    requireNonBlank(condition.range(), "autofilter sort condition range");
    if (condition.color() != null) {
      requireCellColorShape(condition.color(), "autofilter sort color");
    }
    if (condition.iconId() != null) {
      require(condition.iconId() >= 0, "autofilter sort iconId must not be negative");
    }
  }

  private static void requireTableColumnShape(TableColumnReport column) {
    require(column != null, "table column must not be null");
    require(column.id() >= 0L, "table column id must not be negative");
    require(column.name() != null, "table column name must not be null");
  }

  private static void requireCellGradientFillShape(CellGradientFillReport gradient) {
    require(gradient != null, "gradient fill must not be null");
    requireNonBlank(gradient.type(), "gradient fill type");
    require(gradient.stops() != null, "gradient fill stops must not be null");
    require(!gradient.stops().isEmpty(), "gradient fill stops must not be empty");
    for (CellGradientStopReport stop : gradient.stops()) {
      require(stop != null, "gradient fill stop must not be null");
      require(
          Double.isFinite(stop.position()) && stop.position() >= 0.0d && stop.position() <= 1.0d,
          "gradient fill stop position must be between 0.0 and 1.0");
      requireCellColorShape(stop.color(), "gradient fill stop color");
    }
  }

  private static void requireCellColorShape(CellColorReport color, String label) {
    require(color != null, label + " must not be null");
    require(
        color.rgb() != null || color.theme() != null || color.indexed() != null,
        label + " must expose rgb, theme, or indexed semantics");
    if (color.rgb() != null) {
      requireNonBlank(color.rgb(), label + " rgb");
    }
    if (color.theme() != null) {
      require(color.theme() >= 0, label + " theme must not be negative");
    }
    if (color.indexed() != null) {
      require(color.indexed() >= 0, label + " indexed must not be negative");
    }
    if (color.tint() != null) {
      require(Double.isFinite(color.tint()), label + " tint must be finite");
    }
  }

  private static void requireCellProtectionShape(CellProtectionReport protection) {
    require(protection != null, "protection must not be null");
  }

  private static void requireFontHeightShape(FontHeightReport fontHeight) {
    require(fontHeight != null, "fontHeight must not be null");
    ExcelFontHeight expected = new ExcelFontHeight(fontHeight.twips());
    require(
        expected.points().compareTo(fontHeight.points()) == 0,
        "fontHeight points must match twips");
  }

  private static void require(boolean condition, String message) {
    if (!condition) {
      throw new IllegalStateException(message);
    }
  }

  private static void requireNonBlank(String value, String fieldName) {
    require(value != null, fieldName + " must not be null");
    require(!value.isBlank(), fieldName + " must not be blank");
  }
}
