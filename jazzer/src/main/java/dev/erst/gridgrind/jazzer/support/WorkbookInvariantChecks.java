package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.protocol.FontHeightReport;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
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
        require(success.workbook() != null, "workbook summary must not be null");
        require(success.namedRanges() != null, "namedRanges must not be null");
        require(success.sheets() != null, "sheet reports must not be null");
        require(
            success.workbook().sheetCount() == success.workbook().sheetNames().size(),
            "sheetCount must match sheetNames size");
        require(
            success.workbook().namedRangeCount() >= success.namedRanges().size(),
            "namedRangeCount must be greater than or equal to namedRanges size");
        success.namedRanges().forEach(WorkbookInvariantChecks::requireNamedRangeShape);
        success.sheets().forEach(WorkbookInvariantChecks::requireSheetReportShape);
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

  /** Requires the response shape to agree with the request's source and persistence contract. */
  public static void requireWorkflowOutcomeShape(
      GridGrindRequest request, GridGrindResponse response) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(response, "response must not be null");

    if (!(response instanceof GridGrindResponse.Success success)) {
      return;
    }

    switch (request.persistence()) {
      case GridGrindRequest.WorkbookPersistence.None _ ->
          require(success.savedWorkbookPath() == null, "NONE persistence must not return a saved path");
      case GridGrindRequest.WorkbookPersistence.OverwriteSource _ ->
          requireSavedWorkbookPath(success.savedWorkbookPath());
      case GridGrindRequest.WorkbookPersistence.SaveAs _ ->
          requireSavedWorkbookPath(success.savedWorkbookPath());
    }

    switch (request.analysis().namedRanges()) {
      case GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.None _ ->
          require(success.namedRanges().isEmpty(), "NONE named-range analysis must return no named ranges");
      case GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.All _ ->
          require(
              success.workbook().namedRangeCount() == success.namedRanges().size(),
              "ALL named-range analysis must return every named range");
      case GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected _ ->
          require(
              success.workbook().namedRangeCount() >= success.namedRanges().size(),
              "SELECTED named-range analysis must not exceed total named-range count");
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
    workbook.sheetNames().forEach(sheetName -> require(!sheetName.isBlank(), "sheetName must not be blank"));
  }

  private static void requireSavedWorkbookPath(String savedWorkbookPath) {
    require(savedWorkbookPath != null, "savedWorkbookPath must not be null");
    require(savedWorkbookPath.endsWith(".xlsx"), "savedWorkbookPath must point to .xlsx");
    require(Files.exists(Path.of(savedWorkbookPath)), "savedWorkbookPath must exist");
  }

  private static void requireSheetReportShape(GridGrindResponse.SheetReport sheetReport) {
    require(sheetReport.sheetName() != null, "sheetName must not be null");
    require(!sheetReport.sheetName().isBlank(), "sheetName must not be blank");
    require(sheetReport.requestedCells() != null, "requestedCells must not be null");
    require(sheetReport.previewRows() != null, "previewRows must not be null");
    sheetReport.requestedCells().forEach(WorkbookInvariantChecks::requireCellReportShape);
    sheetReport.previewRows()
        .forEach(
            previewRow -> {
              require(previewRow.rowIndex() >= 0, "preview row index must not be negative");
              require(previewRow.cells() != null, "preview row cells must not be null");
              previewRow.cells().forEach(WorkbookInvariantChecks::requireCellReportShape);
            });
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
      require(cellReport.hyperlink().type() != null, "hyperlink type must not be null");
      require(cellReport.hyperlink().target() != null, "hyperlink target must not be null");
      require(!cellReport.hyperlink().target().isBlank(), "hyperlink target must not be blank");
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
        require(!range.target().sheetName().isBlank(), "namedRange target sheet must not be blank");
        require(!range.target().range().isBlank(), "namedRange target range must not be blank");
      }
      case GridGrindResponse.NamedRangeReport.FormulaReport _ -> {}
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
        expected.points().compareTo(fontHeight.points()) == 0, "fontHeight points must match twips");
  }

  private static void require(boolean condition, String message) {
    if (!condition) {
      throw new IllegalStateException(message);
    }
  }
}
