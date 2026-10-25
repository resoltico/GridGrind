package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;

/** Shared selector helpers for shipped example workbook plans. */
final class ExampleSelectors {
  private ExampleSelectors() {}

  static WorkbookSelector.Current workbook() {
    return new WorkbookSelector.Current();
  }

  static SheetSelector.ByName sheet(String name) {
    return new SheetSelector.ByName(new SheetName(name));
  }

  static SheetSelector.ByNames sheets(String... names) {
    return new SheetSelector.ByNames(java.util.Arrays.stream(names).map(SheetName::new).toList());
  }

  static TableSelector.ByNameOnSheet table(String name, String sheetName) {
    return new TableSelector.ByNameOnSheet(name, new SheetName(sheetName));
  }

  static CellSelector.ByAddress cell(String sheetName, String address) {
    return new CellSelector.ByAddress(new SheetName(sheetName), new CellAddress(address));
  }

  static CellSelector.ByAddresses cells(String sheetName, String... addresses) {
    return new CellSelector.ByAddresses(
        new SheetName(sheetName),
        java.util.Arrays.stream(addresses)
            .map(dev.erst.gridgrind.core.vocabulary.CellAddress::new)
            .toList());
  }

  static RangeSelector.ByRange range(String sheetName, String range) {
    return new RangeSelector.ByRange(new SheetName(sheetName), new RangeRef(range));
  }

  static RangeSelector.RectangularWindow window(
      String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new RangeSelector.RectangularWindow(
        new SheetName(sheetName), new CellAddress(topLeftAddress), rowCount, columnCount);
  }
}
