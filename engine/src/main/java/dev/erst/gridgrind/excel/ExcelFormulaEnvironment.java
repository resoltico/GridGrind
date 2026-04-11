package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Workbook-scoped formula-evaluation environment with external workbook bindings and UDFs. */
public record ExcelFormulaEnvironment(
    List<ExcelFormulaExternalWorkbookBinding> externalWorkbooks,
    ExcelFormulaMissingWorkbookPolicy missingWorkbookPolicy,
    List<ExcelFormulaUdfToolpack> udfToolpacks) {
  public ExcelFormulaEnvironment {
    externalWorkbooks = copyExternalWorkbooks(externalWorkbooks);
    missingWorkbookPolicy =
        missingWorkbookPolicy == null
            ? ExcelFormulaMissingWorkbookPolicy.ERROR
            : missingWorkbookPolicy;
    udfToolpacks = copyUdfToolpacks(udfToolpacks);
    requireDistinctUdfFunctions(udfToolpacks);
  }

  /**
   * Returns the default environment with no external bindings, no UDFs, and strict missing refs.
   */
  public static ExcelFormulaEnvironment defaults() {
    return new ExcelFormulaEnvironment(
        List.of(), ExcelFormulaMissingWorkbookPolicy.ERROR, List.of());
  }

  /** Returns whether this environment is the default evaluator state. */
  public boolean isDefault() {
    return externalWorkbooks.isEmpty()
        && missingWorkbookPolicy == ExcelFormulaMissingWorkbookPolicy.ERROR
        && udfToolpacks.isEmpty();
  }

  /** Returns the runtime context snapshot exposed to analysis code. */
  ExcelFormulaRuntimeContext runtimeContext() {
    return new ExcelFormulaRuntimeContext(
        externalWorkbooks.stream()
            .map(ExcelFormulaExternalWorkbookBinding::workbookName)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()),
        missingWorkbookPolicy,
        udfToolpacks.stream()
            .flatMap(toolpack -> toolpack.functions().stream())
            .map(ExcelFormulaUdfFunction::name)
            .collect(java.util.stream.Collectors.toUnmodifiableSet()));
  }

  private static List<ExcelFormulaExternalWorkbookBinding> copyExternalWorkbooks(
      List<ExcelFormulaExternalWorkbookBinding> externalWorkbooks) {
    if (externalWorkbooks == null) {
      return List.of();
    }
    List<ExcelFormulaExternalWorkbookBinding> copy = List.copyOf(externalWorkbooks);
    Set<String> seen = new LinkedHashSet<>();
    for (ExcelFormulaExternalWorkbookBinding externalWorkbook : copy) {
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

  private static List<ExcelFormulaUdfToolpack> copyUdfToolpacks(
      List<ExcelFormulaUdfToolpack> udfToolpacks) {
    if (udfToolpacks == null) {
      return List.of();
    }
    List<ExcelFormulaUdfToolpack> copy = List.copyOf(udfToolpacks);
    Set<String> seen = new LinkedHashSet<>();
    for (ExcelFormulaUdfToolpack toolpack : copy) {
      Objects.requireNonNull(toolpack, "udfToolpacks must not contain nulls");
      String normalized = toolpack.name().toUpperCase(Locale.ROOT);
      if (!seen.add(normalized)) {
        throw new IllegalArgumentException(
            "udfToolpacks must not contain duplicate names: " + toolpack.name());
      }
    }
    return copy;
  }

  private static void requireDistinctUdfFunctions(List<ExcelFormulaUdfToolpack> udfToolpacks) {
    Set<String> seen = new LinkedHashSet<>();
    for (ExcelFormulaUdfToolpack toolpack : udfToolpacks) {
      for (ExcelFormulaUdfFunction function : toolpack.functions()) {
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
