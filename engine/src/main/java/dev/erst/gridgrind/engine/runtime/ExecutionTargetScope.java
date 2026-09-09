package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import java.util.List;
import java.util.Optional;

/** Derives selected-sheet containment only when a selector states every affected sheet exactly. */
final class ExecutionTargetScope {
  private ExecutionTargetScope() {}

  static Optional<List<String>> exactSheetNames(Selector selector) {
    return switch (selector) {
      case SheetSelector.ByName byName -> Optional.of(List.of(byName.name()));
      case SheetSelector.ByNames byNames -> Optional.of(byNames.names());
      case CellSelector.AllUsedInSheet cell -> Optional.of(List.of(cell.sheetName()));
      case CellSelector.ByAddress cell -> Optional.of(List.of(cell.sheetName()));
      case CellSelector.ByAddresses cell -> Optional.of(List.of(cell.sheetName()));
      case RangeSelector.AllOnSheet range -> Optional.of(List.of(range.sheetName()));
      case RangeSelector.ByRange range -> Optional.of(List.of(range.sheetName()));
      case RangeSelector.ByRanges range -> Optional.of(List.of(range.sheetName()));
      case RangeSelector.RectangularWindow range -> Optional.of(List.of(range.sheetName()));
      case RowBandSelector.Span rows -> Optional.of(List.of(rows.sheetName()));
      case RowBandSelector.Insertion rows -> Optional.of(List.of(rows.sheetName()));
      case ColumnBandSelector.Span columns -> Optional.of(List.of(columns.sheetName()));
      case ColumnBandSelector.Insertion columns -> Optional.of(List.of(columns.sheetName()));
      case DrawingObjectSelector.AllOnSheet drawing -> Optional.of(List.of(drawing.sheetName()));
      case DrawingObjectSelector.ByName drawing -> Optional.of(List.of(drawing.sheetName()));
      case ChartSelector.AllOnSheet chart -> Optional.of(List.of(chart.sheetName()));
      case ChartSelector.ByName chart -> Optional.of(List.of(chart.sheetName()));
      case TableSelector.ByNameOnSheet table -> Optional.of(List.of(table.sheetName()));
      case PivotTableSelector.ByNameOnSheet pivot -> Optional.of(List.of(pivot.sheetName()));
      case NamedRangeSelector.SheetScope namedRange -> Optional.of(List.of(namedRange.sheetName()));
      case TableRowSelector.AllRows rows -> exactSheetNames(rows.table());
      case TableRowSelector.ByIndex rows -> exactSheetNames(rows.table());
      case TableRowSelector.ByKeyCell rows -> exactSheetNames(rows.table());
      case TableCellSelector.ByColumnName cell -> exactSheetNames(cell.row());
      case SheetSelector.All _ -> Optional.empty();
      case TableSelector.All _ -> Optional.empty();
      case TableSelector.ByName _ -> Optional.empty();
      case TableSelector.ByNames _ -> Optional.empty();
      case PivotTableSelector.All _ -> Optional.empty();
      case PivotTableSelector.ByName _ -> Optional.empty();
      case PivotTableSelector.ByNames _ -> Optional.empty();
      case NamedRangeSelector.All _ -> Optional.empty();
      case NamedRangeSelector.AnyOf _ -> Optional.empty();
      case NamedRangeSelector.ByName _ -> Optional.empty();
      case NamedRangeSelector.ByNames _ -> Optional.empty();
      case NamedRangeSelector.WorkbookScope _ -> Optional.empty();
      case dev.erst.gridgrind.contract.selector.WorkbookSelector.Current _ -> Optional.empty();
    };
  }
}
