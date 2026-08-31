package dev.erst.gridgrind.architecture.fixture;

import java.util.Objects;
import java.util.function.Consumer;

/** Deliberately benign same-name method reference used to verify POI owner matching. */
public final class ArchitectureNonPoiFormulaMethodFixture {
  private ArchitectureNonPoiFormulaMethodFixture() {}

  /** References a non-POI method whose name resembles a formula write. */
  public static Consumer<String> formulaWriter() {
    return ArchitectureNonPoiFormulaMethodFixture::setCellFormula;
  }

  private static void setCellFormula(String formula) {
    Objects.requireNonNull(formula, "formula must not be null");
  }
}
