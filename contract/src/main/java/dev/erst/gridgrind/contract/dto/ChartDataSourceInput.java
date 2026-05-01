package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;

/** Authored chart data source. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartDataSourceInput.Reference.class, name = "REFERENCE"),
  @JsonSubTypes.Type(value = ChartDataSourceInput.StringLiteral.class, name = "STRING_LITERAL"),
  @JsonSubTypes.Type(value = ChartDataSourceInput.NumericLiteral.class, name = "NUMERIC_LITERAL")
})
public sealed interface ChartDataSourceInput
    permits ChartDataSourceInput.Reference,
        ChartDataSourceInput.StringLiteral,
        ChartDataSourceInput.NumericLiteral {
  /** Workbook-backed chart source formula or defined name. */
  record Reference(String formula) implements ChartDataSourceInput {
    public Reference {
      formula = ChartInput.requireNonBlank(formula, "formula");
    }
  }

  /** Literal string values stored directly in the chart part. */
  record StringLiteral(List<String> values) implements ChartDataSourceInput {
    public StringLiteral {
      values = ChartInput.copyValues(values, "values");
    }
  }

  /** Literal numeric values stored directly in the chart part. */
  record NumericLiteral(List<Double> values) implements ChartDataSourceInput {
    public NumericLiteral {
      values = ChartInput.copyValues(values, "values");
    }
  }
}
