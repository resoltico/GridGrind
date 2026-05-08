package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;

/** Protocol-facing explicit sheet range or formula target for named-range authoring. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = NamedRangeTarget.Range.class, name = "RANGE"),
  @JsonSubTypes.Type(value = NamedRangeTarget.Formula.class, name = "FORMULA")
})
public sealed interface NamedRangeTarget permits NamedRangeTarget.Range, NamedRangeTarget.Formula {
  /** Creates a sheet-local cell or rectangular range target. */
  static Range range(String sheetName, String range) {
    return new Range(sheetName, range);
  }

  /** Creates a formula-defined target that is stored exactly as authored. */
  static Formula formula(String formula) {
    return new Formula(formula);
  }

  /** Sheet-qualified explicit cell or rectangular range target. */
  record Range(String sheetName, String range) implements NamedRangeTarget {
    public Range {
      ExcelSheetNames.requireValid(sheetName, "sheetName");
      Objects.requireNonNull(range, "range must not be null");
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  /** Formula-defined named-range target stored exactly as authored. */
  record Formula(String formula) implements NamedRangeTarget {
    public Formula {
      Objects.requireNonNull(formula, "formula must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
    }
  }
}
