package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;

/** Canonical sheet-query factories kept internal to the Java authoring surface. */
final class SheetQueries {
  private SheetQueries() {}

  static SheetIntrospectionQuery.GetSheetSummary sheetSummary() {
    return new SheetIntrospectionQuery.GetSheetSummary();
  }

  static SheetIntrospectionQuery.GetCells cells() {
    return new SheetIntrospectionQuery.GetCells();
  }

  static SheetIntrospectionQuery.GetWindow window() {
    return new SheetIntrospectionQuery.GetWindow();
  }

  static SheetIntrospectionQuery.GetMergedRegions mergedRegions() {
    return new SheetIntrospectionQuery.GetMergedRegions();
  }

  static SheetIntrospectionQuery.GetHyperlinks hyperlinks() {
    return new SheetIntrospectionQuery.GetHyperlinks();
  }

  static SheetIntrospectionQuery.GetComments comments() {
    return new SheetIntrospectionQuery.GetComments();
  }

  static SheetIntrospectionQuery.GetSheetLayout sheetLayout() {
    return new SheetIntrospectionQuery.GetSheetLayout();
  }

  static SheetIntrospectionQuery.GetPrintLayout printLayout() {
    return new SheetIntrospectionQuery.GetPrintLayout();
  }

  static SheetIntrospectionQuery.GetDataValidations dataValidations() {
    return new SheetIntrospectionQuery.GetDataValidations();
  }

  static SheetIntrospectionQuery.GetConditionalFormatting conditionalFormatting() {
    return new SheetIntrospectionQuery.GetConditionalFormatting();
  }

  static SheetIntrospectionQuery.GetAutofilters autofilters() {
    return new SheetIntrospectionQuery.GetAutofilters();
  }
}
