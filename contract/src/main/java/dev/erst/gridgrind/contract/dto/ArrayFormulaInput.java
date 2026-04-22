package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** Dedicated array-formula input used for single-cell or multi-cell array authoring. */
public record ArrayFormulaInput(TextSourceInput source) {
  public ArrayFormulaInput {
    TextSourceInput normalized = Objects.requireNonNull(source, "source must not be null");
    if (normalized instanceof TextSourceInput.Inline inline) {
      String formula = inline.text();
      if (formula.startsWith("{=") && formula.endsWith("}")) {
        formula = formula.substring(2, formula.length() - 1);
      } else if (formula.startsWith("=")) {
        formula = formula.substring(1);
      }
      if (formula.isBlank()) {
        throw new IllegalArgumentException(
            "source must not be blank after stripping formula syntax");
      }
      normalized = new TextSourceInput.Inline(formula);
    }
    source = normalized;
  }
}
