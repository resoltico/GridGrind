package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.value.CellScalarValue;
import java.util.Objects;

/** Assertions over cell values, display values, formulas, and styles. */
public sealed interface CellAssertion extends Assertion
    permits CellAssertion.CellValue,
        CellAssertion.DisplayValue,
        CellAssertion.FormulaText,
        CellAssertion.CellStyle {

  @Variant(
      id = "EXPECT_CELL_VALUE",
      summary = "Require every selected cell to have the exact effective value.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      },
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record CellValue(CellScalarValue expectedValue) implements CellAssertion {
    public CellValue {
      Objects.requireNonNull(expectedValue, "expectedValue must not be null");
    }
  }

  @Variant(
      id = "EXPECT_DISPLAY_VALUE",
      summary = "Require every selected cell to have the exact formatted display string.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      },
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record DisplayValue(String displayValue) implements CellAssertion {
    public DisplayValue {
      Objects.requireNonNull(displayValue, "displayValue must not be null");
    }
  }

  @Variant(
      id = "EXPECT_FORMULA_TEXT",
      summary = "Require every selected cell to store the exact formula text.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      },
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record FormulaText(String formula) implements CellAssertion {
    public FormulaText {
      formula = AssertionSupport.requireNonBlank(formula, "formula");
    }
  }

  @Variant(
      id = "EXPECT_CELL_STYLE",
      summary = "Require every selected cell to have the exact style snapshot.",
      targetSelectors = {
        CellSelector.ByAddress.class,
        CellSelector.ByAddresses.class,
        TableCellSelector.ByColumnName.class
      },
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record CellStyle(CellStyleReport style) implements CellAssertion {
    public CellStyle {
      Objects.requireNonNull(style, "style must not be null");
    }
  }
}
