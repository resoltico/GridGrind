package dev.erst.gridgrind.contract.query;

import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;

/** Sheet-, cell-, and range-scoped factual inspection queries. */
public sealed interface SheetIntrospectionQuery extends InspectionQuery.Introspection
    permits SheetIntrospectionQuery.GetSheetSummary,
        SheetIntrospectionQuery.GetArrayFormulas,
        SheetIntrospectionQuery.GetCells,
        SheetIntrospectionQuery.GetWindow,
        SheetIntrospectionQuery.GetMergedRegions,
        SheetIntrospectionQuery.GetHyperlinks,
        SheetIntrospectionQuery.GetComments,
        SheetIntrospectionQuery.GetSheetLayout,
        SheetIntrospectionQuery.GetPrintLayout,
        SheetIntrospectionQuery.GetDataValidations,
        SheetIntrospectionQuery.GetConditionalFormatting,
        SheetIntrospectionQuery.GetAutofilters {

  @ProtocolTypeMetadata(
      id = "GET_SHEET_SUMMARY",
      summary = "Return structural summary facts for one sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record GetSheetSummary() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_ARRAY_FORMULAS",
      summary = "Return factual array-formula group metadata for the selected sheets.",
      targetSelectors = {SheetSelector.class})
  record GetArrayFormulas() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_CELLS",
      summary = "Return exact cell snapshots for explicit addresses.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      })
  record GetCells() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_WINDOW",
      summary = "Return a rectangular window of cell snapshots.",
      targetSelectors = {RangeSelector.RectangularWindow.class})
  record GetWindow() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_MERGED_REGIONS",
      summary = "Return the merged regions defined on one sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record GetMergedRegions() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_HYPERLINKS",
      summary = "Return hyperlink metadata for selected cells.",
      targetSelectors = {
        CellSelector.AllUsedInSheet.class,
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      })
  record GetHyperlinks() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_COMMENTS",
      summary = "Return comment metadata for selected cells.",
      targetSelectors = {
        CellSelector.AllUsedInSheet.class,
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      })
  record GetComments() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_SHEET_LAYOUT",
      summary =
          "Return supported sheet-layout metadata for one sheet, including presentation, panes,"
              + " tab color, defaults, ignored errors, and row or column outlineLevel state.",
      targetSelectors = {SheetSelector.ByName.class})
  record GetSheetLayout() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_PRINT_LAYOUT",
      summary = "Return supported print-layout metadata for one sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record GetPrintLayout() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_DATA_VALIDATIONS",
      summary = "Return factual data-validation structures for the selected sheet ranges.",
      targetSelectors = {RangeSelector.class})
  record GetDataValidations() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_CONDITIONAL_FORMATTING",
      summary = "Return factual conditional-formatting blocks for the selected sheet ranges.",
      targetSelectors = {RangeSelector.class})
  record GetConditionalFormatting() implements SheetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_AUTOFILTERS",
      summary = "Return sheet- and table-owned autofilter metadata for one sheet.",
      targetSelectors = {SheetSelector.ByName.class})
  record GetAutofilters() implements SheetIntrospectionQuery {}
}
