package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Reads document-intelligence facts from one sheet wrapper. */
final class ExcelDocumentIntrospector {
  private final ExcelDataValidationController dataValidationController;
  private final ExcelConditionalFormattingController conditionalFormattingController;
  private final ExcelAutofilterController autofilterController;
  private final ExcelTableController tableController;
  private final ExcelDrawingController drawingController;

  ExcelDocumentIntrospector() {
    this(
        new ExcelDataValidationController(),
        new ExcelConditionalFormattingController(),
        new ExcelAutofilterController(),
        new ExcelTableController(),
        new ExcelDrawingController());
  }

  ExcelDocumentIntrospector(
      ExcelDataValidationController dataValidationController,
      ExcelConditionalFormattingController conditionalFormattingController,
      ExcelAutofilterController autofilterController,
      ExcelTableController tableController,
      ExcelDrawingController drawingController) {
    this.dataValidationController =
        Objects.requireNonNull(
            dataValidationController, "dataValidationController must not be null");
    this.conditionalFormattingController =
        Objects.requireNonNull(
            conditionalFormattingController, "conditionalFormattingController must not be null");
    this.autofilterController =
        Objects.requireNonNull(autofilterController, "autofilterController must not be null");
    this.tableController =
        Objects.requireNonNull(tableController, "tableController must not be null");
    this.drawingController =
        Objects.requireNonNull(drawingController, "drawingController must not be null");
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  List<ExcelDataValidationSnapshot> dataValidations(
      ExcelSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return dataValidationController.dataValidations(sheet.xssfSheet(), selection);
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  List<ExcelConditionalFormattingBlockSnapshot> conditionalFormatting(
      ExcelSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return conditionalFormattingController.conditionalFormatting(sheet.xssfSheet(), selection);
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  List<ExcelAutofilterSnapshot> autofilters(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");

    ExcelSheet sheet = workbook.sheet(sheetName);
    List<ExcelAutofilterSnapshot> autofilters = new java.util.ArrayList<>();
    autofilters.addAll(autofilterController.sheetOwnedAutofilters(sheet.xssfSheet()));
    autofilters.addAll(tableController.tableOwnedAutofilters(workbook, sheetName));
    return List.copyOf(autofilters);
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  List<ExcelTableSnapshot> tables(ExcelWorkbook workbook, ExcelTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return tableController.tables(workbook, selection);
  }

  /** Returns factual drawing-object metadata for one sheet. */
  List<ExcelDrawingObjectSnapshot> drawingObjects(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    return drawingController.drawingObjects(workbook.sheet(sheetName).xssfSheet());
  }

  /** Returns factual chart metadata for one sheet. */
  List<ExcelChartSnapshot> charts(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    return drawingController.charts(workbook.sheet(sheetName).xssfSheet());
  }

  /** Returns the extracted binary payload for one existing drawing object on one sheet. */
  ExcelDrawingObjectPayload drawingObjectPayload(
      ExcelWorkbook workbook, String sheetName, String objectName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(objectName, "objectName must not be null");
    return drawingController.drawingObjectPayload(
        workbook.sheet(sheetName).xssfSheet(), objectName);
  }
}
