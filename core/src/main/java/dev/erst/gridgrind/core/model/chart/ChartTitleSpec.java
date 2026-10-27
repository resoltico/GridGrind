package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaInputSecurity;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Requested chart title definition. */
@Union(discriminator = "type", catalog = true, displayOrder = 50)
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_CHART_TITLE_FORMULA_INJECTION
})
public sealed interface ChartTitleSpec
    permits dev.erst.gridgrind.core.model.chart.ChartTitleSpec.None,
        dev.erst.gridgrind.core.model.chart.ChartTitleSpec.Text,
        dev.erst.gridgrind.core.model.chart.ChartTitleSpec.Formula {
  /** Remove any chart or series title. */
  @Variant(id = "NONE", summary = "Remove any chart or series title.")
  record None() implements ChartTitleSpec, ResolvedChartTitle {}

  /** Use a static text title. */
  @Variant(id = "TEXT", summary = "Use one explicit static title.")
  record Text(TextSourceSpec source) implements ChartTitleSpec {
    public Text {
      Objects.requireNonNull(source, "source must not be null");
      if (source instanceof TextSourceSpec.Inline inline && inline.text().isBlank()) {
        throw new IllegalArgumentException("source must not be blank");
      }
    }
  }

  /** Bind the title to one workbook formula. */
  @Variant(
      id = "FORMULA",
      summary = "Bind the chart or series title to one workbook formula that resolves to one cell.")
  record Formula(String formula) implements ChartTitleSpec, ResolvedChartTitle {
    public Formula {
      formula = TextValidation.requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula);
    }
  }
}
