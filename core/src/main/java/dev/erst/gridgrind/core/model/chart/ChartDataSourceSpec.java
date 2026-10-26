package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaInputSecurity;
import java.util.List;

/** Authored chart data source. */
@Union(discriminator = "type", catalog = true, displayOrder = 52)
public sealed interface ChartDataSourceSpec
    permits dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec.Reference,
        dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec.StringLiteral,
        dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec.NumericLiteral {
  /** Workbook-backed chart source formula or defined name. */
  @Variant(id = "REFERENCE", summary = "Workbook-backed chart source formula or defined name.")
  record Reference(String formula) implements ChartDataSourceSpec {
    public Reference {
      formula = ChartSpec.requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula); // LIM-027
    }
  }

  /** Literal string values stored directly in the chart part. */
  @Variant(
      id = "STRING_LITERAL",
      summary = "Literal string chart source stored directly in the chart part.")
  record StringLiteral(List<String> values) implements ChartDataSourceSpec {
    public StringLiteral {
      values = ChartSpec.copyValues(values, "values");
    }
  }

  /** Literal numeric values stored directly in the chart part. */
  @Variant(
      id = "NUMERIC_LITERAL",
      summary = "Literal numeric chart source stored directly in the chart part.")
  record NumericLiteral(List<Double> values) implements ChartDataSourceSpec {
    public NumericLiteral {
      values = ChartSpec.copyValues(values, "values");
    }
  }
}
