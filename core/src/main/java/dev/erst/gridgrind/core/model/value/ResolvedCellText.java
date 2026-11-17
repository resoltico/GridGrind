package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.source.OoxmlTextValidation;

/** Materialized plain cell text, validated before any native cell mutation. */
public record ResolvedCellText(String value) implements ResolvedCellWrite {
  public ResolvedCellText {
    value = CellValueSpec.Validation.requireNonBlank(value, "value");
    value = OoxmlTextValidation.requireXml10CharacterData(value, "value");
    CellTextLimits.requireSupportedLength(value, "value");
  }
}
