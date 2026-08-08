package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.List;
import java.util.Optional;

/** Shared read-result assertions and synthetic fixture builders for executor runtime tests. */
class DefaultGridGrindRequestExecutorReadSupport {
  protected DefaultGridGrindRequestExecutorReadSupport() {}

  static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }

  protected final CellStyleReport requireStyle(CellReport cell) {
    return style(cell);
  }

  static CellStyleReport style(CellReport cell) {
    return cell.style().orElseThrow();
  }

  static String displayValue(CellReport cell) {
    return cell.displayValue().orElseThrow();
  }

  static String textValue(CellReport.TextReport cell) {
    return cell.textValue().orElseThrow();
  }

  static List<RichTextRunReport> runs(CellReport.TextReport cell) {
    return cell.runs().orElseThrow();
  }

  static double numberValue(CellReport.NumberReport cell) {
    return cell.numberValue().orElseThrow();
  }

  static boolean booleanValue(CellReport.BooleanReport cell) {
    return cell.booleanValue().orElseThrow();
  }

  static String errorValue(CellReport.ErrorReport cell) {
    return cell.errorValue().orElseThrow();
  }

  static String formulaText(CellReport.FormulaReport cell) {
    return cell.formula().orElseThrow();
  }

  static CellValueReport evaluation(CellReport.FormulaReport cell) {
    return cell.evaluation().orElseThrow();
  }

  static List<CellReport> windowCells(WindowReport window) {
    return switch (window) {
      case WindowReport.Sparse sparse -> sparse.populatedCells();
      case WindowReport.Dense dense -> {
        List<CellReport> flattened = new ArrayList<>();
        for (WindowRowReport row : dense.rows()) {
          flattened.addAll(row.cells());
        }
        yield List.copyOf(flattened);
      }
    };
  }

  static WindowReport.Dense denseWindow(WindowReport window) {
    return cast(WindowReport.Dense.class, window);
  }

  static ExcelCellReadProjection projection() {
    return new ExcelCellReadProjection(EnumSet.allOf(ExcelCellReadFacet.class));
  }

  static SheetIntrospectionQuery.GetCells allFacetCellsQuery() {
    return new SheetIntrospectionQuery.GetCells(Optional.of(allFacetProjection()));
  }

  static SheetIntrospectionQuery.GetWindow allFacetWindowQuery(boolean includeBlanks) {
    return new SheetIntrospectionQuery.GetWindow(Optional.of(allFacetProjection()), includeBlanks);
  }

  static InspectionSurfaceQuery.GetSheetSchema allFacetSheetSchemaQuery() {
    return new InspectionSurfaceQuery.GetSheetSchema(Optional.of(allFacetProjection()));
  }

  static dev.erst.gridgrind.excel.WorkbookSheetResult.CellsResult excelCellsResult(
      String stepId, String sheetName, List<ExcelCellSnapshot> cells) {
    return new dev.erst.gridgrind.excel.WorkbookSheetResult.CellsResult(
        stepId, sheetName, cells, projection(), false);
  }

  static dev.erst.gridgrind.excel.WorkbookSheetResult.WindowResult excelWindowResult(
      String stepId, dev.erst.gridgrind.excel.WorkbookSheetResult.Window window) {
    return new dev.erst.gridgrind.excel.WorkbookSheetResult.WindowResult(
        stepId, window, projection(), true, false);
  }

  static dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchemaResult excelSheetSchemaResult(
      String stepId, dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchema surface) {
    return new dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchemaResult(
        stepId, surface, projection(), false);
  }

  static CellReport.BlankReport blankReadCell(
      String address, String displayValue, CellStyleReport style) {
    return new CellReport.BlankReport(
        address, Optional.of(displayValue), Optional.of(style), Optional.empty(), Optional.empty());
  }

  static CellReport.TextReport textReadCell(
      String address, String displayValue, CellStyleReport style, String textValue) {
    return new CellReport.TextReport(
        address,
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        Optional.of(textValue),
        Optional.empty());
  }

  static CellReport.NumberReport numberReadCell(
      String address, String displayValue, CellStyleReport style, double numberValue) {
    return new CellReport.NumberReport(
        address,
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        Optional.of(numberValue),
        Optional.empty());
  }

  static CellReport.BooleanReport booleanReadCell(
      String address, String displayValue, CellStyleReport style, boolean booleanValue) {
    return new CellReport.BooleanReport(
        address,
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        Optional.of(booleanValue));
  }

  static CellReport.ErrorReport errorReadCell(
      String address, String displayValue, CellStyleReport style, String errorValue) {
    return new CellReport.ErrorReport(
        address,
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        Optional.of(errorValue));
  }

  static CellReport.FormulaReport formulaReadCell(
      String address,
      String displayValue,
      CellStyleReport style,
      String formula,
      CellValueReport evaluation) {
    return new CellReport.FormulaReport(
        address,
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        Optional.of(formula),
        Optional.of(evaluation));
  }

  static <T extends InspectionResult> T read(
      WorkbookResult.Success success, String stepId, Class<T> type) {
    return ExecutorTestPlanSupport.inspection(success, stepId, type);
  }

  static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new ExcelCellFontSnapshot(
            false,
            false,
            "Aptos",
            ExcelFontHeight.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        ExcelCellFillSnapshot.pattern(dev.erst.gridgrind.excel.foundation.ExcelFillPattern.NONE),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  static ExcelPrintSetupSnapshot defaultPrintSetupSnapshot() {
    return new ExcelPrintSetupSnapshot(
        new ExcelPrintMarginsSnapshot(0.0d, 0.0d, 0.0d, 0.0d, 0.0d, 0.0d),
        false,
        false,
        false,
        0,
        false,
        false,
        0,
        false,
        0,
        List.of(),
        List.of());
  }

  static dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot
      defaultSheetPresentationSnapshot() {
    return new dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot(
        ExcelSheetDisplay.defaults(),
        Optional.empty(),
        ExcelSheetOutlineSummary.defaults(),
        ExcelSheetDefaults.defaults(),
        List.of());
  }

  static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(rgb);
  }

  static ExcelSheetProtectionSettings excelProtectionSettings() {
    return new ExcelSheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  private static CellReadProjection allFacetProjection() {
    return CellReadProjection.of(
        CellReadFacet.VALUE,
        CellReadFacet.STYLE,
        CellReadFacet.FORMAT,
        CellReadFacet.HYPERLINK,
        CellReadFacet.COMMENT,
        CellReadFacet.FORMULA,
        CellReadFacet.RICH_TEXT_RUNS,
        CellReadFacet.TEMPORAL);
  }
}
