package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.List;

/** Shared cell and row literal helpers for shipped example workbook plans. */
final class ExampleCellValues {
  private ExampleCellValues() {}

  @SafeVarargs
  static CellGridSpec rows(CellRowSpec... rows) {
    return new CellGridSpec.Typed(List.of(rows).stream().map(CellRowSpec::toCellInputs).toList());
  }

  static CellRowSpec row(CellValueSpec... cells) {
    return new CellRowSpec.Typed(List.of(cells));
  }

  static CellValueSpec.Text text(String value) {
    return new CellValueSpec.Text(TextSourceSpec.inline(value));
  }

  static CellValueSpec.Text textFile(String path) {
    return new CellValueSpec.Text(TextSourceSpec.utf8File(path));
  }

  static CellValueSpec.Formula formula(String value) {
    return new CellValueSpec.Formula(TextSourceSpec.inline(value));
  }

  static CellValueSpec.NumberValue number(double value) {
    return new CellValueSpec.NumberValue(value);
  }

  static CellValueSpec.BooleanValue bool(boolean value) {
    return new CellValueSpec.BooleanValue(value);
  }
}
