package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
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
    return cell.value()
        .map(dev.erst.gridgrind.core.model.value.CellValue.Text::value)
        .orElseThrow();
  }

  static List<RichTextRunReport> runs(CellReport.TextReport cell) {
    return cell.runs().orElseThrow();
  }

  static double numberValue(CellReport.NumberReport cell) {
    return cell.value()
        .map(dev.erst.gridgrind.core.model.value.CellValue.NumberValue::value)
        .orElseThrow();
  }

  static boolean booleanValue(CellReport.BooleanReport cell) {
    return cell.value()
        .map(dev.erst.gridgrind.core.model.value.CellValue.BooleanValue::value)
        .orElseThrow();
  }

  static String errorValue(CellReport.ErrorReport cell) {
    return cell.value()
        .map(dev.erst.gridgrind.core.model.value.CellValue.ErrorValue::value)
        .orElseThrow();
  }

  static String formulaText(CellReport.FormulaReport cell) {
    return cell.formula().orElseThrow();
  }

  static CellValue evaluation(CellReport.FormulaReport cell) {
    return cell.value().orElseThrow();
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
        new StepId(stepId), sheetName, cells, projection(), false);
  }

  static dev.erst.gridgrind.excel.WorkbookSheetResult.WindowResult excelWindowResult(
      String stepId, dev.erst.gridgrind.excel.WorkbookSheetResult.Window window) {
    return new dev.erst.gridgrind.excel.WorkbookSheetResult.WindowResult(
        new StepId(stepId), window, projection(), true, false);
  }

  static dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchemaResult excelSheetSchemaResult(
      String stepId, dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchema surface) {
    return new dev.erst.gridgrind.excel.WorkbookSurfaceResult.SheetSchemaResult(
        new StepId(stepId), surface, projection(), false);
  }

  static CellReport.BlankReport blankReadCell(
      String address, String displayValue, CellStyleReport style) {
    return new CellReport.BlankReport(
        new CellAddress(address),
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        java.util.Optional.empty());
  }

  static CellReport.TextReport textReadCell(
      String address, String displayValue, CellStyleReport style, String textValue) {
    return new CellReport.TextReport(
        new CellAddress(address),
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        java.util.Optional.of(new dev.erst.gridgrind.core.model.value.CellValue.Text(textValue)),
        Optional.empty());
  }

  static CellReport.NumberReport numberReadCell(
      String address, String displayValue, CellStyleReport style, double numberValue) {
    return new CellReport.NumberReport(
        new CellAddress(address),
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        java.util.Optional.of(
            new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(numberValue)),
        Optional.empty());
  }

  static CellReport.BooleanReport booleanReadCell(
      String address, String displayValue, CellStyleReport style, boolean booleanValue) {
    return new CellReport.BooleanReport(
        new CellAddress(address),
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        java.util.Optional.of(
            new dev.erst.gridgrind.core.model.value.CellValue.BooleanValue(booleanValue)));
  }

  static CellReport.ErrorReport errorReadCell(
      String address, String displayValue, CellStyleReport style, String errorValue) {
    return new CellReport.ErrorReport(
        new CellAddress(address),
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        java.util.Optional.of(
            new dev.erst.gridgrind.core.model.value.CellValue.ErrorValue(errorValue)));
  }

  static CellReport.FormulaReport formulaReadCell(
      String address,
      String displayValue,
      CellStyleReport style,
      String formula,
      CellValue evaluation) {
    return new CellReport.FormulaReport(
        new CellAddress(address),
        Optional.of(displayValue),
        Optional.of(style),
        Optional.empty(),
        Optional.empty(),
        Optional.of(formula),
        Optional.of(evaluation),
        Optional.empty(),
        Optional.empty());
  }

  static <T extends InspectionResult> T read(
      WorkbookResult.Success success, String stepId, Class<T> type) {
    return ExecutorTestPlanSupport.inspection(success, stepId, type);
  }

  static CellStyleReport defaultStyle() {
    return new CellStyleReport(
        "",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            FontHeightReport.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        CellFillReport.pattern(dev.erst.gridgrind.core.vocabulary.ExcelFillPattern.NONE),
        new CellBorderReport(
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None()),
        new CellProtectionReport(true, false));
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
    return CellColorReport.rgb(new RgbHex(rgb));
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
