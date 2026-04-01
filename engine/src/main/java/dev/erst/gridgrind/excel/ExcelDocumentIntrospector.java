package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Reads document-intelligence facts from one sheet wrapper. */
final class ExcelDocumentIntrospector {
  private final ExcelDataValidationController dataValidationController;

  ExcelDocumentIntrospector() {
    this(new ExcelDataValidationController());
  }

  ExcelDocumentIntrospector(ExcelDataValidationController dataValidationController) {
    this.dataValidationController =
        Objects.requireNonNull(
            dataValidationController, "dataValidationController must not be null");
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  List<ExcelDataValidationSnapshot> dataValidations(
      ExcelSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return dataValidationController.dataValidations(sheet.xssfSheet(), selection);
  }
}
