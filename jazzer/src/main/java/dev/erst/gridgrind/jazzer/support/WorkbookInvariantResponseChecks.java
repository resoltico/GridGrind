package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks.require;
import static dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks.requireNonBlank;

import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
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

  static void requireResponseShape(WorkbookResult response) {
    require(response != null, "response must not be null");
    require(response.protocolVersion() != null, "protocolVersion must not be null");

    switch (response) {
      case WorkbookResult.Success success -> requireSuccessResponseShape(success);
      case WorkbookResult.Failure failure -> requireFailureResponseShape(failure);
    }
  }

  static void requireWorkflowOutcomeShape(WorkbookPlan request, WorkbookResult response) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(response, "response must not be null");
    requireResponseShape(response);

    WorkbookPlan.StepPartition stepPartition = request.stepPartition();
    if (response instanceof WorkbookResult.Failure failure) {
      requirePersistenceMatchesRequest(request, failure.persistence());
      require(
          failure.assertions().size() <= stepPartition.assertions().size(),
          "failure assertions size must not exceed the requested assertion count");
      for (int index = 0; index < failure.assertions().size(); index++) {
        requireAssertionMatchesRequest(
            stepPartition.assertions().get(index), failure.assertions().get(index));
      }
      return;
    }

    WorkbookResult.Success success = (WorkbookResult.Success) response;
    requirePersistenceMatchesRequest(request, success.persistence());
    require(
        success.assertions().size() == stepPartition.assertions().size(),
        "assertions size must match the requested assertion count");
    for (int index = 0; index < stepPartition.assertions().size(); index++) {
      requireAssertionMatchesRequest(
          stepPartition.assertions().get(index), success.assertions().get(index));
    }
    require(
        success.inspections().size() == stepPartition.inspections().size(),
        "inspections size must match the requested inspection count");
    for (int index = 0; index < stepPartition.inspections().size(); index++) {
      WorkbookInvariantInspectionResultChecks.requireReadMatchesRequest(
          stepPartition.inspections().get(index), success.inspections().get(index));
    }
  }

  private static void requireSuccessResponseShape(WorkbookResult.Success success) {
    require(success.persistence() != null, "persistence must not be null");
    requirePersistenceOutcomeShape(success.persistence());
    require(success.warnings() != null, "warnings must not be null");
    success.warnings().forEach(WorkbookInvariantResponseChecks::requireRequestWarningShape);
    require(success.assertions() != null, "assertions must not be null");
    success.assertions().forEach(WorkbookInvariantResponseChecks::requireAssertionResultShape);
    require(success.inspections() != null, "inspections must not be null");
    success.inspections().forEach(WorkbookInvariantResponseChecks::requireReadResultShape);
  }

  private static void requireFailureResponseShape(WorkbookResult.Failure failure) {
    require(failure.persistence() != null, "persistence must not be null");
    requirePersistenceOutcomeShape(failure.persistence());
    require(failure.assertions() != null, "assertions must not be null");
    failure.assertions().forEach(WorkbookInvariantResponseChecks::requireAssertionResultShape);
    require(failure.problem() != null, "problem must not be null");
    require(failure.problem().code() != null, "problem code must not be null");
    require(failure.problem().category() != null, "problem category must not be null");
    require(failure.problem().recovery() != null, "problem recovery must not be null");
    require(failure.problem().title() != null, "problem title must not be null");
    require(failure.problem().message() != null, "problem message must not be null");
    require(failure.problem().resolution() != null, "problem resolution must not be null");
    require(failure.problem().context() != null, "problem context must not be null");
    require(failure.problem().causes() != null, "problem causes must not be null");
    if (failure.problem().code()
        == dev.erst.gridgrind.contract.dto.GridGrindProblemCode.ASSERTION_FAILED) {
      require(
          failure.assertions().stream()
              .anyMatch(result -> result instanceof AssertionResult.Failed),
          "ASSERTION_FAILED problems must carry failed assertion evidence in assertions[]");
    }
  }

  private static void requirePersistenceMatchesRequest(
      WorkbookPlan request, WorkbookResultPersistence.PersistenceOutcome persistenceOutcome) {
    WorkbookPlan.WorkbookPersistence requestPersistence = request.persistence();
    switch (requestPersistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> {
        switch (persistenceOutcome) {
          case WorkbookResultPersistence.PersistenceOutcome.NotSaved _ -> {}
          case WorkbookResultPersistence.PersistenceOutcome.SavedAs _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
          case WorkbookResultPersistence.PersistenceOutcome.Overwritten _ ->
              throw new IllegalStateException("NONE persistence must return NONE outcome");
        }
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> {
        requirePersistenceOutcomeShape(persistenceOutcome);
        require(
            persistenceOutcome instanceof WorkbookResultPersistence.PersistenceOutcome.Overwritten,
            "OVERWRITE persistence must return OVERWRITE outcome");
        WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten =
            (WorkbookResultPersistence.PersistenceOutcome.Overwritten) persistenceOutcome;
        switch (request.source()) {
          case WorkbookPlan.WorkbookSource.ExistingFile existingFile -> {
            require(
                overwritten.sourcePath().isPresent(),
                "OVERWRITE persistence with an EXISTING source must include sourcePath");
            require(
                existingFile.path().equals(overwritten.sourcePath().orElseThrow()),
                "OVERWRITE persistence sourcePath must echo the request source path");
          }
          case WorkbookPlan.WorkbookSource.New _ ->
              require(
                  overwritten.sourcePath().isEmpty(),
                  "OVERWRITE persistence with a NEW source must omit sourcePath");
        }
      }
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> {
        requirePersistenceOutcomeShape(persistenceOutcome);
        require(
            persistenceOutcome instanceof WorkbookResultPersistence.PersistenceOutcome.SavedAs,
            "SAVE_AS persistence must return SAVE_AS outcome");
      }
    }
  }

  private static void requirePersistenceOutcomeShape(
      WorkbookResultPersistence.PersistenceOutcome persistenceOutcome) {
    switch (persistenceOutcome) {
      case WorkbookResultPersistence.PersistenceOutcome.NotSaved _ -> {}
      case WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs -> {
        requireNonBlank(savedAs.requestedPath(), "requestedPath");
        requireWriteResultShape(savedAs.write());
      }
      case WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten -> {
        overwritten.sourcePath().ifPresent(sourcePath -> requireNonBlank(sourcePath, "sourcePath"));
        requireWriteResultShape(overwritten.write());
      }
    }
  }

  private static void requireWriteResultShape(WorkbookResultPersistence.WriteResult write) {
    switch (write) {
      case WorkbookResultPersistence.WriteResult.NotWritten _ -> {}
      case WorkbookResultPersistence.WriteResult.Written written ->
          requireExecutionWorkbookPath(written.executionPath());
    }
  }

  private static void requireExecutionWorkbookPath(String executionPath) {
    require(executionPath != null, "executionPath must not be null");
    require(executionPath.endsWith(".xlsx"), "executionPath must point to .xlsx");
    require(Files.exists(Path.of(executionPath)), "executionPath must exist");
  }

  private static void requireRequestWarningShape(RequestWarning warning) {
    require(warning != null, "warning must not be null");
    switch (warning.location()) {
      case RequestWarningLocation.Step step -> {
        require(step.stepIndex() >= 0, "warning stepIndex must not be negative");
        requireNonBlank(step.stepId(), "warning stepId");
        requireNonBlank(step.stepType(), "warning stepType");
      }
      case RequestWarningLocation.RequestPath requestPath -> {
        requireNonBlank(requestPath.path(), "warning request path");
        requireNonBlank(requestPath.pathRole(), "warning request path role");
      }
      case RequestWarningLocation.RequestByteOffset requestByteOffset ->
          require(requestByteOffset.byteOffset() >= 0, "warning byteOffset must not be negative");
      case RequestWarningLocation.FormulaCell formulaCell -> {
        requireNonBlank(formulaCell.sheetName(), "warning formula sheetName");
        requireNonBlank(formulaCell.address(), "warning formula address");
        requireNonBlank(formulaCell.formula(), "warning formula");
      }
    }
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
    switch (assertionResult) {
      case AssertionResult.Passed passed ->
          require(
              passed.outcome() == dev.erst.gridgrind.contract.assertion.AssertionOutcome.PASSED,
              "passed assertion result must report PASSED");
      case AssertionResult.Failed failed -> {
        require(
            failed.outcome() == dev.erst.gridgrind.contract.assertion.AssertionOutcome.FAILED,
            "failed assertion result must report FAILED");
        requireAssertionFailureShape(failed.failure());
        require(
            failed.stepId().equals(failed.failure().stepId()),
            "failed assertion result stepId must match failure evidence");
        require(
            failed.assertionType().equals(failed.failure().assertionType()),
            "failed assertion result type must match failure evidence");
      }
    }
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
      case WorkbookInspectionResult.WorkbookSummaryResult result ->
          requireWorkbookSummaryShape(result.workbook());
      case WorkbookInspectionResult.PackageSecurityResult result ->
          requirePackageSecurityShape(result.security());
      case WorkbookInspectionResult.WorkbookProtectionResult result ->
          requireWorkbookProtectionShape(result.protection());
      case WorkbookInspectionResult.CustomXmlMappingsResult result ->
          result.mappings().forEach(WorkbookInvariantResponseChecks::requireCustomXmlMappingShape);
      case WorkbookInspectionResult.CustomXmlExportResult result ->
          requireCustomXmlExportShape(result.export());
      case WorkbookInspectionResult.NamedRangesResult result ->
          result.namedRanges().forEach(WorkbookInvariantResponseChecks::requireNamedRangeShape);
      case WorkbookInspectionResult.SheetsResult result ->
          result
              .sheetNames()
              .forEach(
                  name ->
                      require(
                          name != null && !name.isBlank(), "sheets sheetName must not be blank"));
      case SheetInspectionResult.SheetSummaryResult result ->
          requireSheetSummaryShape(result.sheet());
      case SheetInspectionResult.ArrayFormulasResult result ->
          result.arrayFormulas().forEach(WorkbookInvariantResponseChecks::requireArrayFormulaShape);
      case SheetInspectionResult.CellsResult result -> {
        require(result.sheetName() != null, "cells sheetName must not be null");
        require(!result.sheetName().isBlank(), "cells sheetName must not be blank");
        result.cells().forEach(WorkbookInvariantResponseChecks::requireCellReportShape);
      }
      case SheetInspectionResult.WindowResult result -> requireWindowShape(result.window());
      case SheetInspectionResult.MergedRegionsResult result -> {
        require(result.sheetName() != null, "merged regions sheetName must not be null");
        require(!result.sheetName().isBlank(), "merged regions sheetName must not be blank");
        result
            .mergedRegions()
            .forEach(
                region ->
                    require(!region.range().isBlank(), "merged region range must not be blank"));
      }
      case SheetInspectionResult.HyperlinksResult result -> {
        require(result.sheetName() != null, "hyperlinks sheetName must not be null");
        require(!result.sheetName().isBlank(), "hyperlinks sheetName must not be blank");
        result.hyperlinks().forEach(WorkbookInvariantResponseChecks::requireHyperlinkEntryShape);
      }
      case SheetInspectionResult.CommentsResult result -> {
        require(result.sheetName() != null, "comments sheetName must not be null");
        require(!result.sheetName().isBlank(), "comments sheetName must not be blank");
        result.comments().forEach(WorkbookInvariantResponseChecks::requireCommentEntryShape);
      }
      case WorkbookAssetInspectionResult.DrawingObjectsResult result -> {
        require(result.sheetName() != null, "drawing objects sheetName must not be null");
        require(!result.sheetName().isBlank(), "drawing objects sheetName must not be blank");
        result.drawingObjects().forEach(WorkbookInvariantResponseChecks::requireDrawingObjectShape);
      }
      case WorkbookAssetInspectionResult.ChartsResult result -> {
        require(result.sheetName() != null, "charts sheetName must not be null");
        require(!result.sheetName().isBlank(), "charts sheetName must not be blank");
        result.charts().forEach(WorkbookInvariantResponseChecks::requireChartReportShape);
      }
      case WorkbookAssetInspectionResult.PivotTablesResult result ->
          result.pivotTables().forEach(WorkbookInvariantResponseChecks::requirePivotTableShape);
      case WorkbookAssetInspectionResult.DrawingObjectPayloadResult result -> {
        require(result.sheetName() != null, "drawing payload sheetName must not be null");
        require(!result.sheetName().isBlank(), "drawing payload sheetName must not be blank");
        requireDrawingObjectPayloadShape(result.payload());
      }
      case SheetInspectionResult.SheetLayoutResult result ->
          requireSheetLayoutShape(result.layout());
      case SheetInspectionResult.PrintLayoutResult result ->
          requirePrintLayoutShape(result.layout());
      case SheetInspectionResult.DataValidationsResult result -> {
        require(result.sheetName() != null, "data validations sheetName must not be null");
        require(!result.sheetName().isBlank(), "data validations sheetName must not be blank");
        result
            .validations()
            .forEach(WorkbookInvariantResponseChecks::requireDataValidationEntryShape);
      }
      case SheetInspectionResult.ConditionalFormattingResult result -> {
        require(result.sheetName() != null, "conditional formatting sheetName must not be null");
        require(
            !result.sheetName().isBlank(), "conditional formatting sheetName must not be blank");
        result
            .conditionalFormattingBlocks()
            .forEach(WorkbookInvariantResponseChecks::requireConditionalFormattingEntryShape);
      }
      case SheetInspectionResult.AutofiltersResult result -> {
        require(result.sheetName() != null, "autofilters sheetName must not be null");
        require(!result.sheetName().isBlank(), "autofilters sheetName must not be blank");
        result.autofilters().forEach(WorkbookInvariantResponseChecks::requireAutofilterEntryShape);
      }
      case WorkbookAssetInspectionResult.TablesResult result ->
          result.tables().forEach(WorkbookInvariantResponseChecks::requireTableEntryShape);
      case WorkbookSurfaceInspectionResult.FormulaSurfaceResult result ->
          requireFormulaSurfaceShape(result.surface());
      case WorkbookSurfaceInspectionResult.SheetSchemaResult result ->
          requireSheetSchemaShape(result.surface());
      case WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult result ->
          requireNamedRangeSurfaceShape(result.surface());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.FormulaHealthResult result ->
          requireFormulaHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.DataValidationHealthResult
              result ->
          requireDataValidationHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult
                  .ConditionalFormattingHealthResult
              result ->
          requireConditionalFormattingHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.AutofilterHealthResult result ->
          requireAutofilterHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.TableHealthResult result ->
          requireTableHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.PivotTableHealthResult result ->
          requirePivotTableHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.HyperlinkHealthResult result ->
          requireHyperlinkHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.NamedRangeHealthResult result ->
          requireNamedRangeHealthShape(result.analysis());
      case dev.erst.gridgrind.contract.query.WorkbookAnalysisResult.WorkbookFindingsResult result ->
          requireWorkbookFindingsShape(result.analysis());
    }
  }

  static void requireWorkbookSummaryShape(WorkbookSummary workbook) {
    WorkbookInvariantWorkbookSurfaceChecks.requireWorkbookSummaryShape(workbook);
  }

  static void requireSheetSummaryShape(SheetSummaryReport sheet) {
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

  private static void requireWindowShape(WindowReport window) {
    WorkbookInvariantAnalysisSurfaceChecks.requireWindowShape(window);
  }

  private static void requireHyperlinkEntryShape(CellHyperlinkReport hyperlink) {
    WorkbookInvariantAnalysisSurfaceChecks.requireHyperlinkEntryShape(hyperlink);
  }

  private static void requireCommentEntryShape(CellCommentReport comment) {
    WorkbookInvariantAnalysisSurfaceChecks.requireCommentEntryShape(comment);
  }

  private static void requireSheetLayoutShape(SheetLayoutReport layout) {
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

  private static void requireFormulaSurfaceShape(FormulaSurfaceReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireFormulaSurfaceShape(analysis);
  }

  private static void requireSheetSchemaShape(SheetSchemaReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireSheetSchemaShape(analysis);
  }

  private static void requireNamedRangeSurfaceShape(NamedRangeSurfaceReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireNamedRangeSurfaceShape(analysis);
  }

  static void requireFormulaHealthShape(FormulaHealthReport analysis) {
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

  static void requireHyperlinkHealthShape(HyperlinkHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireHyperlinkHealthShape(analysis);
  }

  static void requireNamedRangeHealthShape(NamedRangeHealthReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireNamedRangeHealthShape(analysis);
  }

  static void requireWorkbookFindingsShape(WorkbookFindingsReport analysis) {
    WorkbookInvariantAnalysisSurfaceChecks.requireWorkbookFindingsShape(analysis);
  }

  private static void requireCellReportShape(
      dev.erst.gridgrind.contract.dto.CellReport cellReport) {
    WorkbookInvariantCellSurfaceChecks.requireCellReportShape(cellReport);
  }

  static void requireNamedRangeShape(NamedRangeReport namedRange) {
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
