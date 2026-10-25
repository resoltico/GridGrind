package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import java.util.List;
import java.util.Optional;

/** Derives selected-sheet containment only when a selector states every affected sheet exactly. */
final class ExecutionTargetScope {
  private ExecutionTargetScope() {}

  static Optional<List<String>> exactSheetNames(Selector selector) {
    return switch (selector) {
      case SheetSelector.ByName byName -> Optional.of(List.of(byName.name().value()));
      case SheetSelector.ByNames byNames ->
          Optional.of(
              byNames.names().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
                  .toList());
      case CellSelector.AllUsedInSheet cell -> Optional.of(List.of(cell.sheetName().value()));
      case CellSelector.ByAddress cell -> Optional.of(List.of(cell.sheetName().value()));
      case CellSelector.ByAddresses cell -> Optional.of(List.of(cell.sheetName().value()));
      case RangeSelector.AllOnSheet range -> Optional.of(List.of(range.sheetName().value()));
      case RangeSelector.ByRange range -> Optional.of(List.of(range.sheetName().value()));
      case RangeSelector.ByRanges range -> Optional.of(List.of(range.sheetName().value()));
      case RangeSelector.RectangularWindow range -> Optional.of(List.of(range.sheetName().value()));
      case RowBandSelector.Span rows -> Optional.of(List.of(rows.sheetName().value()));
      case RowBandSelector.Insertion rows -> Optional.of(List.of(rows.sheetName().value()));
      case ColumnBandSelector.Span columns -> Optional.of(List.of(columns.sheetName().value()));
      case ColumnBandSelector.Insertion columns ->
          Optional.of(List.of(columns.sheetName().value()));
      case DrawingObjectSelector.AllOnSheet drawing ->
          Optional.of(List.of(drawing.sheetName().value()));
      case DrawingObjectSelector.ByName drawing ->
          Optional.of(List.of(drawing.sheetName().value()));
      case ChartSelector.AllOnSheet chart -> Optional.of(List.of(chart.sheetName().value()));
      case ChartSelector.ByName chart -> Optional.of(List.of(chart.sheetName().value()));
      case TableSelector.ByNameOnSheet table -> Optional.of(List.of(table.sheetName().value()));
      case PivotTableSelector.ByNameOnSheet pivot ->
          Optional.of(List.of(pivot.sheetName().value()));
      case NamedRangeSelector.SheetScope namedRange ->
          Optional.of(List.of(namedRange.sheetName().value()));
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
      case dev.erst.gridgrind.core.plan.WorkbookSelector.Current _ -> Optional.empty();
    };
  }
}
