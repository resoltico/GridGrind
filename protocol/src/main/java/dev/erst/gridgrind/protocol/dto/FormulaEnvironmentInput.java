package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Request-scoped formula-evaluation environment for external workbooks and UDFs. */
public record FormulaEnvironmentInput(
    List<FormulaExternalWorkbookInput> externalWorkbooks,
    FormulaMissingWorkbookPolicy missingWorkbookPolicy,
    List<FormulaUdfToolpackInput> udfToolpacks) {
  public FormulaEnvironmentInput {
    externalWorkbooks = copyExternalWorkbooks(externalWorkbooks);
    missingWorkbookPolicy =
        missingWorkbookPolicy == null ? FormulaMissingWorkbookPolicy.ERROR : missingWorkbookPolicy;
    udfToolpacks = copyUdfToolpacks(udfToolpacks);
    requireDistinctUdfFunctions(udfToolpacks);
  }

  /** Returns whether this environment carries any behavior beyond the default evaluator state. */
  @JsonIgnore
  public boolean isEmpty() {
    return externalWorkbooks.isEmpty()
        && missingWorkbookPolicy == FormulaMissingWorkbookPolicy.ERROR
        && udfToolpacks.isEmpty();
  }

  private static List<FormulaExternalWorkbookInput> copyExternalWorkbooks(
      List<FormulaExternalWorkbookInput> externalWorkbooks) {
    if (externalWorkbooks == null) {
      return List.of();
    }
    List<FormulaExternalWorkbookInput> copy = List.copyOf(externalWorkbooks);
    Set<String> seen = new LinkedHashSet<>();
    for (FormulaExternalWorkbookInput externalWorkbook : copy) {
      Objects.requireNonNull(externalWorkbook, "externalWorkbooks must not contain nulls");
      String normalized = externalWorkbook.workbookName().toUpperCase(Locale.ROOT);
      if (!seen.add(normalized)) {
        throw new IllegalArgumentException(
            "externalWorkbooks must not contain duplicate workbookName values: "
                + externalWorkbook.workbookName());
      }
    }
    return copy;
  }

  private static List<FormulaUdfToolpackInput> copyUdfToolpacks(
      List<FormulaUdfToolpackInput> udfToolpacks) {
    if (udfToolpacks == null) {
      return List.of();
    }
    List<FormulaUdfToolpackInput> copy = List.copyOf(udfToolpacks);
    Set<String> seen = new LinkedHashSet<>();
    for (FormulaUdfToolpackInput udfToolpack : copy) {
      Objects.requireNonNull(udfToolpack, "udfToolpacks must not contain nulls");
      String normalized = udfToolpack.name().toUpperCase(Locale.ROOT);
      if (!seen.add(normalized)) {
        throw new IllegalArgumentException(
            "udfToolpacks must not contain duplicate names: " + udfToolpack.name());
      }
    }
    return copy;
  }

  private static void requireDistinctUdfFunctions(List<FormulaUdfToolpackInput> udfToolpacks) {
    Set<String> seen = new LinkedHashSet<>();
    for (FormulaUdfToolpackInput toolpack : udfToolpacks) {
      for (FormulaUdfFunctionInput function : toolpack.functions()) {
        String normalized = function.name().toUpperCase(Locale.ROOT);
        if (!seen.add(normalized)) {
          throw new IllegalArgumentException(
              "udfToolpacks must not define duplicate function names across toolpacks: "
                  + function.name());
        }
      }
    }
  }
}
