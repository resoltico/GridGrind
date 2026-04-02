package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.protocol.AutofilterEntryReport;
import dev.erst.gridgrind.protocol.AutofilterHealthReport;
import dev.erst.gridgrind.protocol.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.protocol.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.protocol.ConditionalFormattingRuleReport;
import dev.erst.gridgrind.protocol.ConditionalFormattingThresholdReport;
import dev.erst.gridgrind.protocol.DifferentialBorderReport;
import dev.erst.gridgrind.protocol.DifferentialBorderSideReport;
import dev.erst.gridgrind.protocol.DifferentialStyleReport;
import dev.erst.gridgrind.protocol.FontHeightReport;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.HyperlinkTarget;
import dev.erst.gridgrind.protocol.TableEntryReport;
import dev.erst.gridgrind.protocol.TableHealthReport;
import dev.erst.gridgrind.protocol.TableStyleReport;
import dev.erst.gridgrind.protocol.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.WorkbookReadResult;
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

  /** Requires the response shape to agree with the request's source, reads, and persistence contract. */
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
                readExecutor.apply(
                    workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook-shape")).getFirst())
            .workbook();

    requireEngineWorkbookSummaryShape(workbookSummary);
    workbook
        .sheetNames()
        .forEach(
            sheetName ->
                requireEngineSheetSummaryShape(
                    ((dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult)
                            readExecutor.apply(
                                workbook,
                                new WorkbookReadCommand.GetSheetSummary(
                                    "sheet-shape-" + sheetName, sheetName))
                                .getFirst())
                        .sheet()));
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
      case WorkbookReadOperation.GetNamedRanges _ -> {
        WorkbookReadResult.NamedRangesResult result = (WorkbookReadResult.NamedRangesResult) readResult;
        result.namedRanges().forEach(WorkbookInvariantChecks::requireNamedRangeShape);
      }
      case WorkbookReadOperation.GetSheetSummary expected -> {
        WorkbookReadResult.SheetSummaryResult result =
            (WorkbookReadResult.SheetSummaryResult) readResult;
        requireSheetSummaryShape(result.sheet());
        require(expected.sheetName().equals(result.sheet().sheetName()), "sheet summary sheet mismatch");
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
        require(
            expected.rowCount() == result.window().rowCount(),
            "window rowCount mismatch");
        require(
            expected.columnCount() == result.window().columnCount(),
            "window columnCount mismatch");
      }
      case WorkbookReadOperation.GetMergedRegions expected -> {
        WorkbookReadResult.MergedRegionsResult result =
            (WorkbookReadResult.MergedRegionsResult) readResult;
        require(
            expected.sheetName().equals(result.sheetName()), "merged regions sheet mismatch");
      }
      case WorkbookReadOperation.GetHyperlinks expected -> {
        WorkbookReadResult.HyperlinksResult result = (WorkbookReadResult.HyperlinksResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "hyperlinks sheet mismatch");
      }
      case WorkbookReadOperation.GetComments expected -> {
        WorkbookReadResult.CommentsResult result = (WorkbookReadResult.CommentsResult) readResult;
        require(expected.sheetName().equals(result.sheetName()), "comments sheet mismatch");
      }
      case WorkbookReadOperation.GetSheetLayout expected -> {
        WorkbookReadResult.SheetLayoutResult result =
            (WorkbookReadResult.SheetLayoutResult) readResult;
        require(expected.sheetName().equals(result.layout().sheetName()), "layout sheet mismatch");
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
        WorkbookReadResult.SheetSchemaResult result = (WorkbookReadResult.SheetSchemaResult) readResult;
        require(expected.sheetName().equals(result.analysis().sheetName()), "schema sheet mismatch");
        require(
            expected.topLeftAddress().equals(result.analysis().topLeftAddress()),
            "schema topLeftAddress mismatch");
      }
      case WorkbookReadOperation.GetNamedRangeSurface _ -> {
        WorkbookReadResult.NamedRangeSurfaceResult result =
            (WorkbookReadResult.NamedRangeSurfaceResult) readResult;
        require(
            result.analysis().namedRanges() != null, "named range surface entries must not be null");
      }
      case WorkbookReadOperation.AnalyzeFormulaHealth _ ->
          requireFormulaHealthShape(((WorkbookReadResult.FormulaHealthResult) readResult).analysis());
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
      case WorkbookReadResult.NamedRangesResult _ -> "GET_NAMED_RANGES";
      case WorkbookReadResult.SheetSummaryResult _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadResult.CellsResult _ -> "GET_CELLS";
      case WorkbookReadResult.WindowResult _ -> "GET_WINDOW";
      case WorkbookReadResult.MergedRegionsResult _ -> "GET_MERGED_REGIONS";
      case WorkbookReadResult.HyperlinksResult _ -> "GET_HYPERLINKS";
      case WorkbookReadResult.CommentsResult _ -> "GET_COMMENTS";
      case WorkbookReadResult.SheetLayoutResult _ -> "GET_SHEET_LAYOUT";
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
        result.mergedRegions().forEach(region -> require(!region.range().isBlank(), "merged region range must not be blank"));
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
      case WorkbookReadResult.SheetLayoutResult result -> requireSheetLayoutShape(result.layout());
      case WorkbookReadResult.DataValidationsResult result -> {
        require(result.sheetName() != null, "data validations sheetName must not be null");
        require(!result.sheetName().isBlank(), "data validations sheetName must not be blank");
        result.validations().forEach(WorkbookInvariantChecks::requireDataValidationEntryShape);
      }
      case WorkbookReadResult.ConditionalFormattingResult result -> {
        require(
            result.sheetName() != null, "conditional formatting sheetName must not be null");
        require(
            !result.sheetName().isBlank(),
            "conditional formatting sheetName must not be blank");
        result.conditionalFormattingBlocks()
            .forEach(WorkbookInvariantChecks::requireConditionalFormattingEntryShape);
      }
      case WorkbookReadResult.AutofiltersResult result -> {
        require(result.sheetName() != null, "autofilters sheetName must not be null");
        require(!result.sheetName().isBlank(), "autofilters sheetName must not be blank");
        result.autofilters().forEach(WorkbookInvariantChecks::requireAutofilterEntryShape);
      }
      case WorkbookReadResult.TablesResult result ->
          result.tables().forEach(WorkbookInvariantChecks::requireTableEntryShape);
      case WorkbookReadResult.FormulaSurfaceResult result -> requireFormulaSurfaceShape(result.analysis());
      case WorkbookReadResult.SheetSchemaResult result -> requireSheetSchemaShape(result.analysis());
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
      case WorkbookReadResult.TableHealthResult result -> requireTableHealthShape(result.analysis());
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
        .forEach(sheetName -> require(sheetName != null && !sheetName.isBlank(), "sheetName must not be blank"));
    switch (workbook) {
      case GridGrindResponse.WorkbookSummary.Empty empty -> {
        require(empty.sheetCount() == 0, "empty workbook summary must have sheetCount 0");
        require(empty.sheetNames().isEmpty(), "empty workbook summary must have no sheet names");
      }
      case GridGrindResponse.WorkbookSummary.WithSheets withSheets -> {
        require(
            withSheets.sheetCount() > 0, "non-empty workbook summary must have positive sheetCount");
        requireNonBlank(withSheets.activeSheetName(), "activeSheetName");
        require(
            withSheets.sheetNames().contains(withSheets.activeSheetName()),
            "activeSheetName must be present in sheetNames");
        require(
            withSheets.selectedSheetNames() != null,
            "selectedSheetNames must not be null");
        require(
            !withSheets.selectedSheetNames().isEmpty(),
            "selectedSheetNames must not be empty");
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
    workbook
        .sheetNames()
        .forEach(sheetName -> requireNonBlank(sheetName, "engine sheetName"));
    switch (workbook) {
      case dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty empty -> {
        require(empty.sheetCount() == 0, "engine empty workbook summary must have sheetCount 0");
        require(
            empty.sheetNames().isEmpty(),
            "engine empty workbook summary must have no sheet names");
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

  private static void requireWindowShape(GridGrindResponse.WindowReport window) {
    require(window.sheetName() != null, "window sheetName must not be null");
    require(!window.sheetName().isBlank(), "window sheetName must not be blank");
    require(window.topLeftAddress() != null, "window topLeftAddress must not be null");
    require(!window.topLeftAddress().isBlank(), "window topLeftAddress must not be blank");
    require(window.rows() != null, "window rows must not be null");
    require(window.rows().size() == window.rowCount(), "window rows size must match rowCount");
    window.rows()
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
    require(comment.comment().text() != null, "comment text must not be null");
    require(comment.comment().author() != null, "comment author must not be null");
    require(!comment.comment().text().isBlank(), "comment text must not be blank");
    require(!comment.comment().author().isBlank(), "comment author must not be blank");
  }

  private static void requireSheetLayoutShape(GridGrindResponse.SheetLayoutReport layout) {
    require(layout.sheetName() != null, "layout sheetName must not be null");
    require(!layout.sheetName().isBlank(), "layout sheetName must not be blank");
    require(layout.freezePanes() != null, "freezePanes must not be null");
    switch (layout.freezePanes()) {
      case GridGrindResponse.FreezePaneReport.None _ -> {}
      case GridGrindResponse.FreezePaneReport.Frozen frozen -> {
        require(frozen.splitColumn() >= 0, "splitColumn must not be negative");
        require(frozen.splitRow() >= 0, "splitRow must not be negative");
        require(frozen.leftmostColumn() >= 0, "leftmostColumn must not be negative");
        require(frozen.topRow() >= 0, "topRow must not be negative");
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

  private static void requireDataValidationEntryShape(
      dev.erst.gridgrind.protocol.DataValidationEntryReport validation) {
    require(validation.ranges() != null, "data validation ranges must not be null");
    require(!validation.ranges().isEmpty(), "data validation ranges must not be empty");
    validation.ranges().forEach(range -> requireNonBlank(range, "data validation range"));

    switch (validation) {
      case dev.erst.gridgrind.protocol.DataValidationEntryReport.Supported supported ->
          requireSupportedDataValidationShape(supported.validation());
      case dev.erst.gridgrind.protocol.DataValidationEntryReport.Unsupported unsupported -> {
        requireNonBlank(unsupported.kind(), "data validation kind");
        requireNonBlank(unsupported.detail(), "data validation detail");
      }
    }
  }

  private static void requireAutofilterEntryShape(AutofilterEntryReport autofilter) {
    requireNonBlank(autofilter.range(), "autofilter range");
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
    require(
        conditionalFormatting.rules() != null, "conditional formatting rules must not be null");
    require(
        !conditionalFormatting.rules().isEmpty(),
        "conditional formatting rules must not be empty");
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
    table.columnNames().forEach(columnName -> require(columnName != null, "table column name must not be null"));
    require(table.style() != null, "table style must not be null");
    requireTableStyleShape(table.style());
  }

  private static void requireConditionalFormattingRuleShape(
      ConditionalFormattingRuleReport rule) {
    require(rule.priority() > 0, "conditional formatting priority must be greater than 0");
    switch (rule) {
      case ConditionalFormattingRuleReport.FormulaRule formulaRule -> {
        requireNonBlank(formulaRule.formula(), "conditional formatting formula");
        require(formulaRule.style() != null, "conditional formatting style must not be null");
        requireDifferentialStyleShape(formulaRule.style());
      }
      case ConditionalFormattingRuleReport.CellValueRule cellValueRule -> {
        require(cellValueRule.operator() != null, "conditional formatting operator must not be null");
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
            !colorScaleRule.colors().isEmpty(),
            "conditional formatting colors must not be empty");
        colorScaleRule.colors().forEach(color -> requireNonBlank(color, "conditional formatting color"));
      }
      case ConditionalFormattingRuleReport.DataBarRule dataBarRule -> {
        requireNonBlank(dataBarRule.color(), "conditional formatting color");
        requireConditionalFormattingThresholdShape(dataBarRule.minThreshold());
        requireConditionalFormattingThresholdShape(dataBarRule.maxThreshold());
        require(dataBarRule.widthMin() >= 0, "conditional formatting widthMin must not be negative");
        require(dataBarRule.widthMax() >= 0, "conditional formatting widthMax must not be negative");
      }
      case ConditionalFormattingRuleReport.IconSetRule iconSetRule -> {
        require(iconSetRule.iconSet() != null, "conditional formatting iconSet must not be null");
        require(iconSetRule.thresholds() != null, "conditional formatting thresholds must not be null");
        require(
            !iconSetRule.thresholds().isEmpty(),
            "conditional formatting thresholds must not be empty");
        iconSetRule
            .thresholds()
            .forEach(WorkbookInvariantChecks::requireConditionalFormattingThresholdShape);
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
    require(style.unsupportedFeatures() != null, "conditional formatting unsupportedFeatures must not be null");
    style
        .unsupportedFeatures()
        .forEach(feature -> require(feature != null, "conditional formatting unsupported feature must not be null"));
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
      dev.erst.gridgrind.protocol.DataValidationEntryReport.DataValidationDefinitionReport
          validation) {
    require(validation != null, "data validation definition must not be null");
    require(validation.rule() != null, "data validation rule must not be null");
    switch (validation.rule()) {
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.ExplicitList explicitList -> {
        require(explicitList.values() != null, "explicit list values must not be null");
        require(!explicitList.values().isEmpty(), "explicit list values must not be empty");
        explicitList.values().forEach(value -> requireNonBlank(value, "explicit list value"));
      }
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.FormulaList formulaList ->
          requireNonBlank(formulaList.formula(), "formula list formula");
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.WholeNumber wholeNumber ->
          requireComparisonRuleShape(wholeNumber.operator(), wholeNumber.formula1());
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.DecimalNumber decimalNumber ->
          requireComparisonRuleShape(decimalNumber.operator(), decimalNumber.formula1());
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.DateRule dateRule ->
          requireComparisonRuleShape(dateRule.operator(), dateRule.formula1());
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.TimeRule timeRule ->
          requireComparisonRuleShape(timeRule.operator(), timeRule.formula1());
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.TextLength textLength ->
          requireComparisonRuleShape(textLength.operator(), textLength.formula1());
      case dev.erst.gridgrind.protocol.DataValidationRuleInput.CustomFormula customFormula ->
          requireNonBlank(customFormula.formula(), "custom validation formula");
    }
    if (validation.prompt() != null) {
      requireNonBlank(validation.prompt().title(), "data validation prompt title");
      requireNonBlank(validation.prompt().text(), "data validation prompt text");
    }
    if (validation.errorAlert() != null) {
      require(validation.errorAlert().style() != null, "data validation error style must not be null");
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
                  sheet.distinctFormulaCount() >= 0,
                  "distinctFormulaCount must not be negative");
              sheet
                  .formulas()
                  .forEach(
                      formula -> {
                        require(formula.formula() != null, "formula pattern must not be null");
                        require(
                            !formula.formula().isBlank(),
                            "formula pattern must not be blank");
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
    require(
        analysis.workbookScopedCount() >= 0, "workbookScopedCount must not be negative");
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
        analysis.checkedFormulaCellCount() >= 0,
        "checkedFormulaCellCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireDataValidationHealthShape(
      dev.erst.gridgrind.protocol.DataValidationHealthReport analysis) {
    require(
        analysis.checkedValidationCount() >= 0,
        "checkedValidationCount must not be negative");
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
    require(
        analysis.checkedAutofilterCount() >= 0,
        "checkedAutofilterCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireTableHealthShape(TableHealthReport analysis) {
    require(analysis.checkedTableCount() >= 0, "checkedTableCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireHyperlinkHealthShape(
      GridGrindResponse.HyperlinkHealthReport analysis) {
    require(
        analysis.checkedHyperlinkCount() >= 0,
        "checkedHyperlinkCount must not be negative");
    requireAnalysisSummaryShape(analysis.summary(), analysis.findings());
  }

  private static void requireNamedRangeHealthShape(
      GridGrindResponse.NamedRangeHealthReport analysis) {
    require(
        analysis.checkedNamedRangeCount() >= 0,
        "checkedNamedRangeCount must not be negative");
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
        summary.totalCount() == findings.size(),
        "analysis totalCount must match findings size");
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
      case GridGrindResponse.CellReport.TextReport text ->
          require(text.stringValue() != null, "stringValue must not be null");
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
      require(cellReport.comment().text() != null, "comment text must not be null");
      require(cellReport.comment().author() != null, "comment author must not be null");
      require(!cellReport.comment().text().isBlank(), "comment text must not be blank");
      require(!cellReport.comment().author().isBlank(), "comment author must not be blank");
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
        require(
            !range.target().sheetName().isBlank(),
            "namedRange target sheet must not be blank");
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
    require(style.horizontalAlignment() != null, "horizontalAlignment must not be null");
    require(style.verticalAlignment() != null, "verticalAlignment must not be null");
    require(style.fontName() != null, "fontName must not be null");
    require(!style.fontName().isBlank(), "fontName must not be blank");
    requireFontHeightShape(style.fontHeight());
    require(style.topBorderStyle() != null, "topBorderStyle must not be null");
    require(style.rightBorderStyle() != null, "rightBorderStyle must not be null");
    require(style.bottomBorderStyle() != null, "bottomBorderStyle must not be null");
    require(style.leftBorderStyle() != null, "leftBorderStyle must not be null");
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
