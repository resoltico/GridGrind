package dev.erst.gridgrind.core.model.chart;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaInputSecurity;
import dev.erst.gridgrind.core.source.TextSourceInput;
import java.util.Objects;

/** Requested chart title definition. */
@Union(discriminator = "type", catalog = true, displayOrder = 50)
public sealed interface ChartTitleInput
    permits dev.erst.gridgrind.core.model.chart.ChartTitleInput.None,
        dev.erst.gridgrind.core.model.chart.ChartTitleInput.Text,
        dev.erst.gridgrind.core.model.chart.ChartTitleInput.Formula {
  /** Remove any chart or series title. */
  @Variant(id = "NONE", summary = "Remove any chart or series title.")
  record None() implements ChartTitleInput {}

  /** Use a static text title. */
  @Variant(id = "TEXT", summary = "Use one explicit static title.")
  record Text(TextSourceInput source) implements ChartTitleInput {
    public Text {
      Objects.requireNonNull(source, "source must not be null");
      if (source instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
        throw new IllegalArgumentException("source must not be blank");
      }
    }
  }

  /** Bind the title to one workbook formula. */
  @Variant(
      id = "FORMULA",
      summary = "Bind the chart or series title to one workbook formula that resolves to one cell.")
  record Formula(String formula) implements ChartTitleInput {
    public Formula {
      formula = ChartInput.requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula); // LIM-033
    }
  }
}
