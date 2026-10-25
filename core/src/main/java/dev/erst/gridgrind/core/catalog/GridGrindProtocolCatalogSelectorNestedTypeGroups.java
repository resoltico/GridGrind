package dev.erst.gridgrind.core.catalog;

import static dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;
import static dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.selectorDescriptor;

import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import java.util.ArrayList;
import java.util.List;

/** Owns one focused subset of nested-type group descriptors for the protocol catalog. */
final class GridGrindProtocolCatalogSelectorNestedTypeGroups {
  private GridGrindProtocolCatalogSelectorNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> SELECTOR_GROUPS =
      List.of(
          nestedTypeGroup(
              "workbookSelectorTypes",
              dev.erst.gridgrind.core.plan.WorkbookSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.WorkbookSelector.Current.class,
                      "Target the workbook currently being executed."))),
          nestedTypeGroup(
              "sheetSelectorTypes",
              dev.erst.gridgrind.core.plan.SheetSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.SheetSelector.All.class,
                      "Select every sheet in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.SheetSelector.ByName.class,
                      "Select one exact sheet by name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.SheetSelector.ByNames.class,
                      "Select one or more exact sheets by ordered name list."))),
          nestedTypeGroup(
              "cellSelectorTypes",
              dev.erst.gridgrind.core.plan.CellSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.CellSelector.AllUsedInSheet.class,
                      "Select every physically present cell on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.CellSelector.ByAddress.class,
                      "Select one exact cell on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.CellSelector.ByAddresses.class,
                      "Select one or more exact cells on one sheet."))),
          nestedTypeGroup(
              "rangeSelectorTypes",
              dev.erst.gridgrind.core.plan.RangeSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet.class,
                      "Select every matching range-backed structure on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.RangeSelector.ByRange.class,
                      "Select one exact rectangular range on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.RangeSelector.ByRanges.class,
                      "Select one or more exact rectangular ranges on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow.class,
                      "Select one rectangular window anchored at one top-left cell."))),
          nestedTypeGroup(
              "rowBandSelectorTypes",
              dev.erst.gridgrind.core.plan.RowBandSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.RowBandSelector.Span.class,
                      "Select one inclusive zero-based row span on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.RowBandSelector.Insertion.class,
                      "Select one row insertion point plus row count on one sheet."))),
          nestedTypeGroup(
              "columnBandSelectorTypes",
              dev.erst.gridgrind.core.plan.ColumnBandSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.ColumnBandSelector.Span.class,
                      "Select one inclusive zero-based column span on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.ColumnBandSelector.Insertion.class,
                      "Select one column insertion point plus column count on one sheet."))),
          nestedTypeGroup(
              "tableSelectorTypes",
              dev.erst.gridgrind.core.plan.TableSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableSelector.All.class,
                      "Select every table in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableSelector.ByName.class,
                      "Select one workbook-global table by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableSelector.ByNames.class,
                      "Select one or more workbook-global tables by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableSelector.ByNameOnSheet.class,
                      "Select one workbook-global table by exact name and expected owning sheet."))),
          nestedTypeGroup(
              "tableRowSelectorTypes",
              dev.erst.gridgrind.core.plan.TableRowSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableRowSelector.AllRows.class,
                      "Select every logical data row in one selected table."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableRowSelector.ByIndex.class,
                      "Select one zero-based data row by index in one selected table."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableRowSelector.ByKeyCell.class,
                      "Select one logical data row by matching one key-column cell value."))),
          nestedTypeGroup(
              "tableCellSelectorTypes",
              dev.erst.gridgrind.core.plan.TableCellSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.TableCellSelector.ByColumnName.class,
                      "Select one logical cell within one selected table row by column name."))),
          nestedTypeGroup(
              "namedRangeRefSelectorTypes",
              dev.erst.gridgrind.core.plan.NamedRangeSelector.Ref.class,
              namedRangeScopeDescriptors(
                  dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName.class,
                  "Match a named range reference across all scopes by exact name.",
                  dev.erst.gridgrind.core.plan.NamedRangeSelector.WorkbookScope.class,
                  "Match one workbook-scoped named range reference by exact name.",
                  dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope.class,
                  "Match one sheet-scoped named range reference on one sheet.")),
          nestedTypeGroup(
              "namedRangeScopeTypes",
              NamedRangeScope.class,
              List.of(
                  descriptor(NamedRangeScope.Workbook.class, "WORKBOOK", "Target workbook scope."),
                  descriptor(
                      NamedRangeScope.Sheet.class, "SHEET", "Target one specific sheet scope."))),
          nestedTypeGroup(
              "namedRangeSelectorTypes",
              dev.erst.gridgrind.core.plan.NamedRangeSelector.class,
              namedRangeSelectionDescriptors()),
          nestedTypeGroup(
              "drawingObjectSelectorTypes",
              dev.erst.gridgrind.core.plan.DrawingObjectSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet.class,
                      "Select every drawing object on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.DrawingObjectSelector.ByName.class,
                      "Select one drawing object by exact sheet-local object name."))),
          nestedTypeGroup(
              "chartSelectorTypes",
              dev.erst.gridgrind.core.plan.ChartSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.ChartSelector.AllOnSheet.class,
                      "Select every chart on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.ChartSelector.ByName.class,
                      "Select one chart by exact sheet-local chart name."))),
          nestedTypeGroup(
              "pivotTableSelectorTypes",
              dev.erst.gridgrind.core.plan.PivotTableSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.PivotTableSelector.All.class,
                      "Select every pivot table in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.PivotTableSelector.ByName.class,
                      "Select one workbook-global pivot table by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.PivotTableSelector.ByNames.class,
                      "Select one or more workbook-global pivot tables by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.core.plan.PivotTableSelector.ByNameOnSheet.class,
                      "Select one workbook-global pivot table by exact name and expected owning sheet."))));

  private static List<CatalogTypeDescriptor> namedRangeSelectionDescriptors() {
    List<CatalogTypeDescriptor> descriptors = new ArrayList<>();
    descriptors.add(
        selectorDescriptor(
            dev.erst.gridgrind.core.plan.NamedRangeSelector.All.class,
            "Select every user-facing named range."));
    descriptors.add(
        selectorDescriptor(
            dev.erst.gridgrind.core.plan.NamedRangeSelector.AnyOf.class,
            "Select the union of one or more explicit named-range references."));
    descriptors.add(
        selectorDescriptor(
            dev.erst.gridgrind.core.plan.NamedRangeSelector.ByNames.class,
            "Match named ranges across all scopes by exact name set."));
    descriptors.addAll(
        namedRangeScopeDescriptors(
            dev.erst.gridgrind.core.plan.NamedRangeSelector.ByName.class,
            "Match a named range across all scopes by exact name.",
            dev.erst.gridgrind.core.plan.NamedRangeSelector.WorkbookScope.class,
            "Match the workbook-scoped named range with the exact name.",
            dev.erst.gridgrind.core.plan.NamedRangeSelector.SheetScope.class,
            "Match the sheet-scoped named range on one sheet."));
    return List.copyOf(descriptors);
  }

  private static <
          B extends Record & Selector, W extends Record & Selector, S extends Record & Selector>
      List<CatalogTypeDescriptor> namedRangeScopeDescriptors(
          Class<B> byNameType,
          String byNameSummary,
          Class<W> workbookScopeType,
          String workbookScopeSummary,
          Class<S> sheetScopeType,
          String sheetScopeSummary) {
    return List.of(
        selectorDescriptor(byNameType, byNameSummary),
        selectorDescriptor(workbookScopeType, workbookScopeSummary),
        selectorDescriptor(sheetScopeType, sheetScopeSummary));
  }
}
