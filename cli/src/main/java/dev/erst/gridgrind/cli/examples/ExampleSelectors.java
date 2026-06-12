package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.List;

/** Shared selector helpers for shipped example workbook plans. */
final class ExampleSelectors {
  private ExampleSelectors() {}

  static WorkbookSelector.Current workbook() {
    return new WorkbookSelector.Current();
  }

  static SheetSelector.ByName sheet(String name) {
    return new SheetSelector.ByName(name);
  }

  static SheetSelector.ByNames sheets(String... names) {
    return new SheetSelector.ByNames(List.of(names));
  }

  static TableSelector.ByNameOnSheet table(String name, String sheetName) {
    return new TableSelector.ByNameOnSheet(name, sheetName);
  }

  static CellSelector.ByAddress cell(String sheetName, String address) {
    return new CellSelector.ByAddress(sheetName, address);
  }

  static CellSelector.ByAddresses cells(String sheetName, String... addresses) {
    return new CellSelector.ByAddresses(sheetName, List.of(addresses));
  }

  static RangeSelector.ByRange range(String sheetName, String range) {
    return new RangeSelector.ByRange(sheetName, range);
  }

  static RangeSelector.RectangularWindow window(
      String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new RangeSelector.RectangularWindow(sheetName, topLeftAddress, rowCount, columnCount);
  }
}
