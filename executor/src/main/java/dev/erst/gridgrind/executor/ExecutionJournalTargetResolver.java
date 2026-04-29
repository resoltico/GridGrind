package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
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
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.ArrayList;
import java.util.List;

/** Canonical target summaries recorded inside execution-journal step entries. */
final class ExecutionJournalTargetResolver {
  private ExecutionJournalTargetResolver() {}

  static List<ExecutionJournal.Target> resolve(WorkbookStep step, ExecutionJournalLevel level) {
    return switch (level) {
      case SUMMARY -> List.of(summaryTarget(step.target()));
      case NORMAL, VERBOSE -> expandedTargets(step.target());
    };
  }

  static ExecutionJournal.Target summaryTarget(Selector selector) {
    return new ExecutionJournal.Target(selectorKind(selector), summaryLabel(selector));
  }

  static List<ExecutionJournal.Target> expandedTargets(Selector selector) {
    if (selector instanceof SheetSelector.ByNames byNames) {
      return byNames.names().stream()
          .map(name -> new ExecutionJournal.Target("SHEET", "Sheet " + name))
          .toList();
    }
    if (selector instanceof CellSelector.ByAddresses byAddresses) {
      return byAddresses.addresses().stream()
          .map(
              address ->
                  new ExecutionJournal.Target(
                      "CELL", "Cell " + byAddresses.sheetName() + "!" + address))
          .toList();
    }
    if (selector instanceof CellSelector.ByQualifiedAddresses qualifiedAddresses) {
      return qualifiedAddresses.cells().stream()
          .map(
              cell ->
                  new ExecutionJournal.Target(
                      "CELL", "Cell " + cell.sheetName() + "!" + cell.address()))
          .toList();
    }
    if (selector instanceof RangeSelector.ByRanges byRanges) {
      return byRanges.ranges().stream()
          .map(
              range ->
                  new ExecutionJournal.Target(
                      "RANGE", "Range " + byRanges.sheetName() + "!" + range))
          .toList();
    }
    if (selector instanceof TableSelector.ByNames byNames) {
      return byNames.names().stream()
          .map(name -> new ExecutionJournal.Target("TABLE", "Table " + name))
          .toList();
    }
    if (selector instanceof PivotTableSelector.ByNames byNames) {
      return byNames.names().stream()
          .map(name -> new ExecutionJournal.Target("PIVOT_TABLE", "Pivot table " + name))
          .toList();
    }
    if (selector instanceof NamedRangeSelector.ByNames byNames) {
      return byNames.names().stream()
          .map(name -> new ExecutionJournal.Target("NAMED_RANGE", "Named range " + name))
          .toList();
    }
    if (selector instanceof NamedRangeSelector.AnyOf anyOf) {
      List<ExecutionJournal.Target> expanded = new ArrayList<>();
      for (NamedRangeSelector.Ref ref : anyOf.selectors()) {
        expanded.addAll(expandedNamedRangeRefTargets(ref));
      }
      return List.copyOf(expanded);
    }
    return List.of(summaryTarget(selector));
  }

  private static List<ExecutionJournal.Target> expandedNamedRangeRefTargets(
      NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName ->
          List.of(new ExecutionJournal.Target("NAMED_RANGE", "Named range " + byName.name()));
      case NamedRangeSelector.WorkbookScope workbookScope ->
          List.of(
              new ExecutionJournal.Target(
                  "NAMED_RANGE", "Workbook-scoped named range " + workbookScope.name()));
      case NamedRangeSelector.SheetScope sheetScope ->
          List.of(
              new ExecutionJournal.Target(
                  "NAMED_RANGE",
                  "Sheet-scoped named range " + sheetScope.sheetName() + "!" + sheetScope.name()));
    };
  }

  private static String selectorKind(Selector selector) {
    return switch (selector) {
      case WorkbookSelector _ -> "WORKBOOK";
      case SheetSelector _ -> "SHEET";
      case CellSelector _ -> "CELL";
      case RangeSelector _ -> "RANGE";
      case RowBandSelector _ -> "ROW_BAND";
      case ColumnBandSelector _ -> "COLUMN_BAND";
      case DrawingObjectSelector _ -> "DRAWING";
      case ChartSelector _ -> "CHART";
      case TableSelector _ -> "TABLE";
      case PivotTableSelector _ -> "PIVOT_TABLE";
      case NamedRangeSelector _ -> "NAMED_RANGE";
      case TableRowSelector _ -> "TABLE_ROW";
      case TableCellSelector _ -> "TABLE_CELL";
    };
  }

  private static String summaryLabel(Selector selector) {
    return switch (selector) {
      case WorkbookSelector.Current _ -> "Current workbook";
      case SheetSelector.All _ -> "All sheets";
      case SheetSelector.ByName byName -> "Sheet " + byName.name();
      case SheetSelector.ByNames byNames -> "Sheets " + byNames.names();
      case CellSelector.AllUsedInSheet allUsedInSheet ->
          "All used cells on " + allUsedInSheet.sheetName();
      case CellSelector.ByAddress byAddress ->
          "Cell " + byAddress.sheetName() + "!" + byAddress.address();
      case CellSelector.ByAddresses byAddresses ->
          "Cells " + byAddresses.sheetName() + "!" + byAddresses.addresses();
      case CellSelector.ByQualifiedAddresses qualifiedAddresses ->
          "Qualified cells " + qualifiedAddresses.cells();
      case RangeSelector.AllOnSheet allOnSheet -> "All ranges on " + allOnSheet.sheetName();
      case RangeSelector.ByRange byRange -> "Range " + byRange.sheetName() + "!" + byRange.range();
      case RangeSelector.ByRanges byRanges ->
          "Ranges " + byRanges.sheetName() + "!" + byRanges.ranges();
      case RangeSelector.RectangularWindow window ->
          "Window " + window.sheetName() + "!" + window.range();
      case RowBandSelector.Span span ->
          "Rows " + span.sheetName() + "[" + span.firstRowIndex() + ":" + span.lastRowIndex() + "]";
      case RowBandSelector.Insertion insertion ->
          "Row insertion "
              + insertion.sheetName()
              + "[before="
              + insertion.beforeRowIndex()
              + ", count="
              + insertion.rowCount()
              + "]";
      case ColumnBandSelector.Span span ->
          "Columns "
              + span.sheetName()
              + "["
              + span.firstColumnIndex()
              + ":"
              + span.lastColumnIndex()
              + "]";
      case ColumnBandSelector.Insertion insertion ->
          "Column insertion "
              + insertion.sheetName()
              + "[before="
              + insertion.beforeColumnIndex()
              + ", count="
              + insertion.columnCount()
              + "]";
      case DrawingObjectSelector.AllOnSheet allOnSheet ->
          "All drawing objects on " + allOnSheet.sheetName();
      case DrawingObjectSelector.ByName byName ->
          "Drawing object " + byName.sheetName() + "!" + byName.objectName();
      case ChartSelector.AllOnSheet allOnSheet -> "All charts on " + allOnSheet.sheetName();
      case ChartSelector.ByName byName -> "Chart " + byName.sheetName() + "!" + byName.chartName();
      case TableSelector.All _ -> "All tables";
      case TableSelector.ByName byName -> "Table " + byName.name();
      case TableSelector.ByNames byNames -> "Tables " + byNames.names();
      case TableSelector.ByNameOnSheet byNameOnSheet ->
          "Table " + byNameOnSheet.sheetName() + "!" + byNameOnSheet.name();
      case PivotTableSelector.All _ -> "All pivot tables";
      case PivotTableSelector.ByName byName -> "Pivot table " + byName.name();
      case PivotTableSelector.ByNames byNames -> "Pivot tables " + byNames.names();
      case PivotTableSelector.ByNameOnSheet byNameOnSheet ->
          "Pivot table " + byNameOnSheet.sheetName() + "!" + byNameOnSheet.name();
      case NamedRangeSelector.All _ -> "All named ranges";
      case NamedRangeSelector.ByName byName -> "Named range " + byName.name();
      case NamedRangeSelector.ByNames byNames -> "Named ranges " + byNames.names();
      case NamedRangeSelector.WorkbookScope workbookScope ->
          "Workbook-scoped named range " + workbookScope.name();
      case NamedRangeSelector.SheetScope sheetScope ->
          "Sheet-scoped named range " + sheetScope.sheetName() + "!" + sheetScope.name();
      case NamedRangeSelector.AnyOf anyOf -> "Named range selector " + anyOf.selectors();
      case TableRowSelector.AllRows allRows -> "All rows in " + summaryLabel(allRows.table());
      case TableRowSelector.ByIndex byIndex ->
          "Row " + byIndex.rowIndex() + " in " + summaryLabel(byIndex.table());
      case TableRowSelector.ByKeyCell byKeyCell ->
          "Row where "
              + byKeyCell.columnName()
              + "="
              + summarizeCellInput(byKeyCell.expectedValue())
              + " in "
              + summaryLabel(byKeyCell.table());
      case TableCellSelector.ByColumnName byColumnName ->
          "Column " + byColumnName.columnName() + " in " + summaryLabel(byColumnName.row());
    };
  }

  private static String summarizeCellInput(CellInput input) {
    if (input instanceof CellInput.Blank) {
      return "Blank[]";
    }
    if (input instanceof CellInput.Text text) {
      return "Text[" + summarizeTextSource(text.source()) + "]";
    }
    if (input instanceof CellInput.Numeric numeric) {
      return "Number[number=" + numeric.number() + "]";
    }
    if (input instanceof CellInput.BooleanValue booleanValue) {
      return "Boolean[value=" + booleanValue.bool() + "]";
    }
    CellInput.Formula formula = (CellInput.Formula) input;
    return "Formula[" + summarizeTextSource(formula.source()) + "]";
  }

  private static String summarizeTextSource(TextSourceInput source) {
    return switch (source) {
      case TextSourceInput.Inline inline -> "text=" + inline.text();
      case TextSourceInput.Utf8File file -> "path=" + file.path();
      case TextSourceInput.StandardInput _ -> "source=STANDARD_INPUT";
    };
  }
}
