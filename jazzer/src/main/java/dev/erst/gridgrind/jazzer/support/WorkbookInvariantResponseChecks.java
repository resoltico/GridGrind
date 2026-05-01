package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks.require;
import static dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks.requireNonBlank;

import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.GridGrindAnalysisReports;
import dev.erst.gridgrind.contract.dto.GridGrindLayoutSurfaceReports;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.GridGrindSchemaAndFormulaReports;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.*;
import dev.erst.gridgrind.contract.step.*;
import dev.erst.gridgrind.excel.*;
import java.nio.file.*;
import java.util.*;

/** Validates response-shape invariants over protocol-level workbook workflow outcomes. */
final class WorkbookInvariantResponseChecks {
  private WorkbookInvariantResponseChecks() {}

  static void requireResponseShape(GridGrindResponse response) {
    require(response != null, "response must not be null");
    require(response.protocolVersion() != null, "protocolVersion must not be null");

    switch (response) {
      case GridGrindResponse.Success success -> requireSuccessResponseShape(success);
      case GridGrindResponse.Failure failure -> requireFailureResponseShape(failure);
    }
  }

  static void requireWorkflowOutcomeShape(WorkbookPlan request, GridGrindResponse response) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(response, "response must not be null");

    switch (response) {
      case GridGrindResponse.Failure _ -> {
        return;
      }
      case GridGrindResponse.Success success -> {
        requirePersistenceMatchesRequest(request.persistence(), success.persistence());
        require(
            success.assertions().size() == request.stepPartition().assertions().size(),
            "assertions size must match the requested assertion count");
        for (int index = 0; index < request.stepPartition().assertions().size(); index++) {
          requireAssertionMatchesRequest(
              request.stepPartition().assertions().get(index), success.assertions().get(index));
        }
        require(
            success.inspections().size() == request.stepPartition().inspections().size(),
            "inspections size must match the requested inspection count");
        for (int index = 0; index < request.stepPartition().inspections().size(); index++) {
          WorkbookInvariantInspectionResultChecks.requireReadMatchesRequest(
              request.stepPartition().inspections().get(index), success.inspections().get(index));
        }
      }
    }
  }

  private static void requireSuccessResponseShape(GridGrindResponse.Success success) {
    require(success.persistence() != null, "persistence must not be null");
    requirePersistenceOutcomeShape(success.persistence());
    require(success.warnings() != null, "warnings must not be null");
    success.warnings().forEach(WorkbookInvariantResponseChecks::requireRequestWarningShape);
    require(success.assertions() != null, "assertions must not be null");
    success.assertions().forEach(WorkbookInvariantResponseChecks::requireAssertionResultShape);
    require(success.inspections() != null, "inspections must not be null");
    success.inspections().forEach(WorkbookInvariantResponseChecks::requireReadResultShape);
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
      GridGrindResponsePersistence.PersistenceOutcome persistenceOutcome) {
    switch (requestPersistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> {
        switch (persistenceOutcome) {
          case GridGrindResponsePersistence.PersistenceOutcome.NotSaved _ -> {}
          case GridGrindResponsePersistence.PersistenceOutcome.SavedAs _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
          case GridGrindResponsePersistence.PersistenceOutcome.Overwritten _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
        }
      }
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> {
        requirePersistenceOutcomeShape(persistenceOutcome);
        require(
            persistenceOutcome
                instanceof GridGrindResponsePersistence.PersistenceOutcome.Overwritten,
            "OVERWRITE persistence must return OVERWRITE outcome");
      }
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> {
        requirePersistenceOutcomeShape(persistenceOutcome);
        require(
            persistenceOutcome instanceof GridGrindResponsePersistence.PersistenceOutcome.SavedAs,
            "SAVE_AS persistence must return SAVE_AS outcome");
      }
    }
  }

  private static void requirePersistenceOutcomeShape(
      GridGrindResponsePersistence.PersistenceOutcome persistenceOutcome) {
    switch (persistenceOutcome) {
      case GridGrindResponsePersistence.PersistenceOutcome.NotSaved _ -> {}
      case GridGrindResponsePersistence.PersistenceOutcome.SavedAs savedAs -> {
        requireNonBlank(savedAs.requestedPath(), "requestedPath");
        requireExecutionWorkbookPath(savedAs.executionPath());
      }
      case GridGrindResponsePersistence.PersistenceOutcome.Overwritten overwritten -> {
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
    assertionFailure
        .observations()
        .forEach(WorkbookInvariantResponseChecks::requireReadResultShape);
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
          result.mappings().forEach(WorkbookInvariantResponseChecks::requireCustomXmlMappingShape);
      case InspectionResult.CustomXmlExportResult result ->
          requireCustomXmlExportShape(result.export());
      case InspectionResult.NamedRangesResult result ->
          result.namedRanges().forEach(WorkbookInvariantResponseChecks::requireNamedRangeShape);
      case InspectionResult.SheetSummaryResult result -> requireSheetSummaryShape(result.sheet());
      case InspectionResult.ArrayFormulasResult result ->
          result.arrayFormulas().forEach(WorkbookInvariantResponseChecks::requireArrayFormulaShape);
      case InspectionResult.CellsResult result -> {
        require(result.sheetName() != null, "cells sheetName must not be null");
        require(!result.sheetName().isBlank(), "cells sheetName must not be blank");
        result.cells().forEach(WorkbookInvariantResponseChecks::requireCellReportShape);
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
        result.hyperlinks().forEach(WorkbookInvariantResponseChecks::requireHyperlinkEntryShape);
      }
      case InspectionResult.CommentsResult result -> {
        require(result.sheetName() != null, "comments sheetName must not be null");
        require(!result.sheetName().isBlank(), "comments sheetName must not be blank");
        result.comments().forEach(WorkbookInvariantResponseChecks::requireCommentEntryShape);
      }
      case InspectionResult.DrawingObjectsResult result -> {
        require(result.sheetName() != null, "drawing objects sheetName must not be null");
        require(!result.sheetName().isBlank(), "drawing objects sheetName must not be blank");
        result.drawingObjects().forEach(WorkbookInvariantResponseChecks::requireDrawingObjectShape);
      }
      case InspectionResult.ChartsResult result -> {
        require(result.sheetName() != null, "charts sheetName must not be null");
        require(!result.sheetName().isBlank(), "charts sheetName must not be blank");
        result.charts().forEach(WorkbookInvariantResponseChecks::requireChartReportShape);
      }
      case InspectionResult.PivotTablesResult result ->
          result.pivotTables().forEach(WorkbookInvariantResponseChecks::requirePivotTableShape);
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
        result
            .validations()
            .forEach(WorkbookInvariantResponseChecks::requireDataValidationEntryShape);
      }
      case InspectionResult.ConditionalFormattingResult result -> {
        require(result.sheetName() != null, "conditional formatting sheetName must not be null");
        require(
            !result.sheetName().isBlank(), "conditional formatting sheetName must not be blank");
        result
            .conditionalFormattingBlocks()
            .forEach(WorkbookInvariantResponseChecks::requireConditionalFormattingEntryShape);
      }
      case InspectionResult.AutofiltersResult result -> {
        require(result.sheetName() != null, "autofilters sheetName must not be null");
        require(!result.sheetName().isBlank(), "autofilters sheetName must not be blank");
        result.autofilters().forEach(WorkbookInvariantResponseChecks::requireAutofilterEntryShape);
      }
      case InspectionResult.TablesResult result ->
          result.tables().forEach(WorkbookInvariantResponseChecks::requireTableEntryShape);
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

  static void requireWorkbookSummaryShape(
      GridGrindWorkbookSurfaceReports.WorkbookSummary workbook) {
    WorkbookInvariantWorkbookSurfaceChecks.requireWorkbookSummaryShape(workbook);
  }

  static void requireSheetSummaryShape(GridGrindWorkbookSurfaceReports.SheetSummaryReport sheet) {
    WorkbookInvariantWorkbookSurfaceChecks.requireSheetSummaryShape(sheet);
  }

  static void requireDrawingObjectShape(DrawingObjectReport drawingObject) {
    WorkbookInvariantWorkbookSurfaceChecks.requireDrawingObjectShape(drawingObject);
  }

  static void requireDrawingObjectPayloadShape(DrawingObjectPayloadReport payload) {
    WorkbookInvariantWorkbookSurfaceChecks.requireDrawingObjectPayloadShape(payload);
  }

  static void requireChartReportShape(ChartReport chart) {
    WorkbookInvariantWorkbookSurfaceChecks.requireChartReportShape(chart);
  }

  private static void requireWindowShape(GridGrindLayoutSurfaceReports.WindowReport window) {
    WorkbookInvariantAnalysisSurfaceChecks.requireWindowShape(window);
  }

  private static void requireHyperlinkEntryShape(
      GridGrindLayoutSurfaceReports.CellHyperlinkReport hyperlink) {
    WorkbookInvariantAnalysisSurfaceChecks.requireHyperlinkEntryShape(hyperlink);
  }

  private static void requireCommentEntryShape(
      GridGrindLayoutSurfaceReports.CellCommentReport comment) {
    WorkbookInvariantAnalysisSurfaceChecks.requireCommentEntryShape(comment);
  }

  private static void requireSheetLayoutShape(
      GridGrindLayoutSurfaceReports.SheetLayoutReport layout) {
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

  static void requireTableEntryShape(TableEntryReport table) {
    WorkbookInvariantAnalysisSurfaceChecks.requireTableEntryShape(table);
  }

  static void requirePivotTableShape(PivotTableReport pivotTable) {
    WorkbookInvariantAnalysisSurfaceChecks.requirePivotTableShape(pivotTable);
  }

  private static void requireFormulaSurfaceShape(
      GridGrindSchemaAndFormulaReports.FormulaSurfaceReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireFormulaSurfaceShape(analysis);
  }

  private static void requireSheetSchemaShape(
      GridGrindSchemaAndFormulaReports.SheetSchemaReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireSheetSchemaShape(analysis);
  }

  private static void requireNamedRangeSurfaceShape(
      GridGrindSchemaAndFormulaReports.NamedRangeSurfaceReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireNamedRangeSurfaceShape(analysis);
  }

  static void requireFormulaHealthShape(GridGrindAnalysisReports.FormulaHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireFormulaHealthShape(analysis);
  }

  static void requireDataValidationHealthShape(
      dev.erst.gridgrind.contract.dto.DataValidationHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireDataValidationHealthShape(analysis);
  }

  static void requireConditionalFormattingHealthShape(ConditionalFormattingHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireConditionalFormattingHealthShape(analysis);
  }

  static void requireAutofilterHealthShape(AutofilterHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireAutofilterHealthShape(analysis);
  }

  static void requireTableHealthShape(TableHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireTableHealthShape(analysis);
  }

  static void requirePivotTableHealthShape(PivotTableHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requirePivotTableHealthShape(analysis);
  }

  static void requireHyperlinkHealthShape(GridGrindAnalysisReports.HyperlinkHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireHyperlinkHealthShape(analysis);
  }

  static void requireNamedRangeHealthShape(
      GridGrindAnalysisReports.NamedRangeHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireNamedRangeHealthShape(analysis);
  }

  static void requireWorkbookFindingsShape(
      GridGrindAnalysisReports.WorkbookFindingsReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireWorkbookFindingsShape(analysis);
  }

  private static void requireCellReportShape(
      dev.erst.gridgrind.contract.dto.CellReport cellReport) {
    WorkbookInvariantCellSurfaceChecks.requireCellReportShape(cellReport);
  }

  static void requireNamedRangeShape(GridGrindWorkbookSurfaceReports.NamedRangeReport namedRange) {
    WorkbookInvariantCellSurfaceChecks.requireNamedRangeShape(namedRange);
  }

  static void requireWorkbookProtectionShape(WorkbookProtectionReport protection) {
    WorkbookInvariantCellSurfaceChecks.requireWorkbookProtectionShape(protection);
  }

  static void requirePackageSecurityShape(OoxmlPackageSecurityReport security) {
    WorkbookInvariantCellSurfaceChecks.requirePackageSecurityShape(security);
  }

  static void requireCustomXmlMappingShape(CustomXmlMappingReport mapping) {
    WorkbookInvariantCellSurfaceChecks.requireCustomXmlMappingShape(mapping);
  }

  static void requireCustomXmlExportShape(CustomXmlExportReport export) {
    WorkbookInvariantCellSurfaceChecks.requireCustomXmlExportShape(export);
  }

  static void requireArrayFormulaShape(ArrayFormulaReport arrayFormula) {
    WorkbookInvariantCellSurfaceChecks.requireArrayFormulaShape(arrayFormula);
  }
}
