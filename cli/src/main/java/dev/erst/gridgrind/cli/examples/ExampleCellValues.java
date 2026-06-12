package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.dto.CellGridInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellRowInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;

/** Shared cell and row literal helpers for shipped example workbook plans. */
final class ExampleCellValues {
  private ExampleCellValues() {}

  @SafeVarargs
  static CellGridInput rows(CellRowInput... rows) {
    return new CellGridInput.Typed(List.of(rows).stream().map(CellRowInput::toCellInputs).toList());
  }

  static CellRowInput row(CellInput... cells) {
    return new CellRowInput.Typed(List.of(cells));
  }

  static CellInput.Text text(String value) {
    return new CellInput.Text(TextSourceInput.inline(value));
  }

  static CellInput.Formula formula(String value) {
    return new CellInput.Formula(TextSourceInput.inline(value));
  }

  static CellInput.NumberValue number(double value) {
    return new CellInput.NumberValue(value);
  }

  static CellInput.BooleanValue bool(boolean value) {
    return new CellInput.BooleanValue(value);
  }
}
