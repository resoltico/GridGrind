package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.selectorDescriptor;

import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import java.util.List;

/** Owns one focused subset of nested-type group descriptors for the protocol catalog. */
final class GridGrindProtocolCatalogSelectorNestedTypeGroups {
  private GridGrindProtocolCatalogSelectorNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> SELECTOR_GROUPS =
      List.of(
          nestedTypeGroup(
              "workbookSelectorTypes",
              dev.erst.gridgrind.contract.selector.WorkbookSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.WorkbookSelector.Current.class,
                      "Target the workbook currently being executed."))),
          nestedTypeGroup(
              "sheetSelectorTypes",
              dev.erst.gridgrind.contract.selector.SheetSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.All.class,
                      "Select every sheet in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class,
                      "Select one exact sheet by name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByNames.class,
                      "Select one or more exact sheets by ordered name list."))),
          nestedTypeGroup(
              "cellSelectorTypes",
              dev.erst.gridgrind.contract.selector.CellSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.AllUsedInSheet.class,
                      "Select every physically present cell on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddress.class,
                      "Select one exact cell on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddresses.class,
                      "Select one or more exact cells on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByQualifiedAddresses.class,
                      "Select one or more exact cells across one or more sheets."))),
          nestedTypeGroup(
              "rangeSelectorTypes",
              dev.erst.gridgrind.contract.selector.RangeSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.AllOnSheet.class,
                      "Select every matching range-backed structure on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRange.class,
                      "Select one exact rectangular range on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRanges.class,
                      "Select one or more exact rectangular ranges on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.RectangularWindow.class,
                      "Select one rectangular window anchored at one top-left cell."))),
          nestedTypeGroup(
              "rowBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.RowBandSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Span.class,
                      "Select one inclusive zero-based row span on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Insertion.class,
                      "Select one row insertion point plus row count on one sheet."))),
          nestedTypeGroup(
              "columnBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.ColumnBandSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Span.class,
                      "Select one inclusive zero-based column span on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Insertion.class,
                      "Select one column insertion point plus column count on one sheet."))),
          nestedTypeGroup(
              "tableSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.All.class,
                      "Select every table in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByName.class,
                      "Select one workbook-global table by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNames.class,
                      "Select one or more workbook-global tables by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet.class,
                      "Select one workbook-global table by exact name and expected owning sheet."))),
          nestedTypeGroup(
              "tableRowSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableRowSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.AllRows.class,
                      "Select every logical data row in one selected table."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByIndex.class,
                      "Select one zero-based data row by index in one selected table."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell.class,
                      "Select one logical data row by matching one key-column cell value."))),
          nestedTypeGroup(
              "tableCellSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableCellSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableCellSelector.ByColumnName.class,
                      "Select one logical cell within one selected table row by column name."))),
          nestedTypeGroup(
              "namedRangeRefSelectorTypes",
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.Ref.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "Match a named range reference across all scopes by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "Match one workbook-scoped named range reference by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "Match one sheet-scoped named range reference on one sheet."))),
          nestedTypeGroup(
              "namedRangeScopeTypes",
              NamedRangeScope.class,
              List.of(
                  descriptor(NamedRangeScope.Workbook.class, "WORKBOOK", "Target workbook scope."),
                  descriptor(
                      NamedRangeScope.Sheet.class, "SHEET", "Target one specific sheet scope."))),
          nestedTypeGroup(
              "namedRangeSelectorTypes",
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.All.class,
                      "Select every user-facing named range."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf.class,
                      "Select the union of one or more explicit named-range references."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "Match a named range across all scopes by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByNames.class,
                      "Match named ranges across all scopes by exact name set."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "Match the workbook-scoped named range with the exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "Match the sheet-scoped named range on one sheet."))),
          nestedTypeGroup(
              "drawingObjectSelectorTypes",
              dev.erst.gridgrind.contract.selector.DrawingObjectSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet.class,
                      "Select every drawing object on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.ByName.class,
                      "Select one drawing object by exact sheet-local object name."))),
          nestedTypeGroup(
              "chartSelectorTypes",
              dev.erst.gridgrind.contract.selector.ChartSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.AllOnSheet.class,
                      "Select every chart on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.ByName.class,
                      "Select one chart by exact sheet-local chart name."))),
          nestedTypeGroup(
              "pivotTableSelectorTypes",
              dev.erst.gridgrind.contract.selector.PivotTableSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.All.class,
                      "Select every pivot table in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByName.class,
                      "Select one workbook-global pivot table by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames.class,
                      "Select one or more workbook-global pivot tables by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNameOnSheet.class,
                      "Select one workbook-global pivot table by exact name and expected owning sheet."))));
}
