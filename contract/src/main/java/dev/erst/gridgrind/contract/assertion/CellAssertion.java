package dev.erst.gridgrind.contract.assertion;

import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import java.util.Objects;

/** Assertions over cell values, display values, formulas, and styles. */
public sealed interface CellAssertion extends Assertion
    permits CellAssertion.CellValue,
        CellAssertion.DisplayValue,
        CellAssertion.FormulaText,
        CellAssertion.CellStyle {

  @ProtocolTypeMetadata(
      id = "EXPECT_CELL_VALUE",
      summary = "Require every selected cell to have the exact effective value.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      })
  record CellValue(CellScalarValue expectedValue) implements CellAssertion {
    public CellValue {
      Objects.requireNonNull(expectedValue, "expectedValue must not be null");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_DISPLAY_VALUE",
      summary = "Require every selected cell to have the exact formatted display string.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      })
  record DisplayValue(String displayValue) implements CellAssertion {
    public DisplayValue {
      Objects.requireNonNull(displayValue, "displayValue must not be null");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_FORMULA_TEXT",
      summary = "Require every selected cell to store the exact formula text.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      })
  record FormulaText(String formula) implements CellAssertion {
    public FormulaText {
      formula = AssertionSupport.requireNonBlank(formula, "formula");
    }
  }

  @ProtocolTypeMetadata(
      id = "EXPECT_CELL_STYLE",
      summary = "Require every selected cell to have the exact style snapshot.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      })
  record CellStyle(CellStyleReport style) implements CellAssertion {
    public CellStyle {
      Objects.requireNonNull(style, "style must not be null");
    }
  }
}
