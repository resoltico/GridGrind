package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Selector builders for the fluent Java authoring API. */
public final class Targets {
  private Targets() {}

  /** Returns one workbook-scoped fluent target. */
  public static WorkbookTarget workbook() {
    return new WorkbookTarget(new WorkbookSelector.Current());
  }

  /** Returns one exact sheet target by sheet name. */
  public static SheetTarget sheet(SheetName name) {
    return new SheetTarget(new SheetSelector.ByName(name));
  }

  /** Returns one exact cell target by sheet name and A1 address. */
  public static CellTarget cell(SheetName sheetName, String address) {
    return new CellTarget(new CellSelector.ByAddress(sheetName, address));
  }

  /** Returns one exact range target by sheet name and A1 range. */
  public static RangeTarget range(SheetName sheetName, String range) {
    return new RangeTarget(new RangeSelector.ByRange(sheetName, range));
  }

  /** Returns one rectangular window target from a top-left cell plus dimensions. */
  public static WindowTarget window(
      SheetName sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new WindowTarget(
        new RangeSelector.RectangularWindow(sheetName, topLeftAddress, rowCount, columnCount));
  }

  /** Returns one exact table target by table name. */
  public static TableTarget table(String name) {
    return new TableTarget(new TableSelector.ByName(name));
  }

  /** Returns one exact table target by table name plus sheet ownership. */
  public static TableTarget tableOnSheet(String name, SheetName sheetName) {
    return new TableTarget(new TableSelector.ByNameOnSheet(name, sheetName));
  }

  /** Returns one exact named-range target by name. */
  public static NamedRangeTarget namedRange(String name) {
    return new NamedRangeTarget(new NamedRangeSelector.ByName(name));
  }

  /** Returns one exact sheet-scoped named-range target. */
  public static NamedRangeTarget namedRangeOnSheet(String name, SheetName sheetName) {
    return new NamedRangeTarget(new NamedRangeSelector.SheetScope(name, sheetName));
  }

  /** Returns one exact workbook-scoped named-range target. */
  public static NamedRangeTarget workbookNamedRange(String name) {
    return new NamedRangeTarget(new NamedRangeSelector.WorkbookScope(name));
  }

  /** Returns one exact chart target by sheet name and chart name. */
  public static ChartTarget chart(SheetName sheetName, String chartName) {
    return new ChartTarget(new ChartSelector.ByName(sheetName, chartName));
  }

  /** Returns one exact pivot-table target by pivot-table name. */
  public static PivotTableTarget pivotTable(String name) {
    return new PivotTableTarget(new PivotTableSelector.ByName(name));
  }

  /** Returns one exact pivot-table target by pivot-table name plus sheet ownership. */
  public static PivotTableTarget pivotTableOnSheet(String name, SheetName sheetName) {
    return new PivotTableTarget(new PivotTableSelector.ByNameOnSheet(name, sheetName));
  }
}
