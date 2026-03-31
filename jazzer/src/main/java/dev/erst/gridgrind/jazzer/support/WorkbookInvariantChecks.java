package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.protocol.FontHeightReport;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.HyperlinkTarget;
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
    require(workbook.sheetCount() >= 0, "sheetCount must not be negative");
    require(
        workbook.sheetCount() == workbook.sheetNames().size(),
        "sheet count must match sheetNames size");
    require(
        workbook.sheetNames().size() == new HashSet<>(workbook.sheetNames()).size(),
        "sheet names must be unique");
    workbook
        .sheetNames()
        .forEach(sheetName -> require(!sheetName.isBlank(), "sheetName must not be blank"));
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
      case WorkbookReadResult.FormulaSurfaceResult _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadResult.SheetSchemaResult _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadResult.NamedRangeSurfaceResult _ -> "GET_NAMED_RANGE_SURFACE";
      case WorkbookReadResult.FormulaHealthResult _ -> "ANALYZE_FORMULA_HEALTH";
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
      case WorkbookReadResult.FormulaSurfaceResult result -> requireFormulaSurfaceShape(result.analysis());
      case WorkbookReadResult.SheetSchemaResult result -> requireSheetSchemaShape(result.analysis());
      case WorkbookReadResult.NamedRangeSurfaceResult result ->
          requireNamedRangeSurfaceShape(result.analysis());
      case WorkbookReadResult.FormulaHealthResult result ->
          requireFormulaHealthShape(result.analysis());
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
  }

  private static void requireSheetSummaryShape(GridGrindResponse.SheetSummaryReport sheet) {
    require(sheet.sheetName() != null, "sheetName must not be null");
    require(!sheet.sheetName().isBlank(), "sheetName must not be blank");
    require(sheet.physicalRowCount() >= 0, "physicalRowCount must not be negative");
    require(sheet.lastRowIndex() >= -1, "lastRowIndex must be greater than or equal to -1");
    require(sheet.lastColumnIndex() >= -1, "lastColumnIndex must be greater than or equal to -1");
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
