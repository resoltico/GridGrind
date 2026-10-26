package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.FormulaExternalWorkbookSpec;
import dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionSpec;
import dev.erst.gridgrind.core.model.formula.FormulaUdfToolpackSpec;
import dev.erst.gridgrind.excel.ExcelFormulaEnvironment;
import dev.erst.gridgrind.excel.ExcelFormulaExternalWorkbookBinding;
import dev.erst.gridgrind.excel.ExcelFormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.excel.ExcelFormulaUdfFunction;
import dev.erst.gridgrind.excel.ExcelFormulaUdfToolpack;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.jspecify.annotations.Nullable;

/** Converts request-scoped formula environment payloads into engine-owned evaluator models. */
final class FormulaEnvironmentConverter {
  private FormulaEnvironmentConverter() {}

  static ExcelFormulaEnvironment toExcelFormulaEnvironment(
      @Nullable FormulaEnvironmentSpec input, ExecutionInputBindings bindings) throws IOException {
    if (input == null) {
      return ExcelFormulaEnvironment.defaults();
    }
    return new ExcelFormulaEnvironment(
        externalWorkbookBindings(input, bindings),
        toExcelMissingWorkbookPolicy(input.missingWorkbookPolicy()),
        input.udfToolpacks().stream()
            .map(FormulaEnvironmentConverter::toExcelUdfToolpack)
            .toList());
  }

  private static List<ExcelFormulaExternalWorkbookBinding> externalWorkbookBindings(
      FormulaEnvironmentSpec input, ExecutionInputBindings bindings) throws IOException {
    List<ExcelFormulaExternalWorkbookBinding> bindingsByName = new ArrayList<>();
    for (FormulaExternalWorkbookSpec externalWorkbook : input.externalWorkbooks()) {
      bindingsByName.add(toExcelFormulaExternalWorkbookBinding(externalWorkbook, bindings));
    }
    return List.copyOf(bindingsByName);
  }

  private static ExcelFormulaExternalWorkbookBinding toExcelFormulaExternalWorkbookBinding(
      FormulaExternalWorkbookSpec input, ExecutionInputBindings bindings) throws IOException {
    return new ExcelFormulaExternalWorkbookBinding(
        input.workbookName(),
        bindings
            .requestPathAccess()
            .materializeRead(
                input.path(),
                "formulaEnvironment.externalWorkbooks",
                "gridgrind-formula-workbook-",
                ".xlsx"));
  }

  private static ExcelFormulaMissingWorkbookPolicy toExcelMissingWorkbookPolicy(
      dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy policy) {
    if (policy == dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy.ERROR) {
      return ExcelFormulaMissingWorkbookPolicy.ERROR;
    }
    return ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE;
  }

  private static ExcelFormulaUdfToolpack toExcelUdfToolpack(FormulaUdfToolpackSpec input) {
    return new ExcelFormulaUdfToolpack(
        input.name(),
        input.functions().stream().map(FormulaEnvironmentConverter::toExcelUdfFunction).toList());
  }

  private static ExcelFormulaUdfFunction toExcelUdfFunction(FormulaUdfFunctionSpec input) {
    return new ExcelFormulaUdfFunction(
        input.name(),
        input.minimumArgumentCount(),
        input.maximumArgumentCount(),
        input.formulaTemplate());
  }
}
