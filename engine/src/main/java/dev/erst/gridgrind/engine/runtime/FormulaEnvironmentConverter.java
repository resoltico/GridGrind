package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
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
      @Nullable FormulaEnvironmentInput input, ExecutionInputBindings bindings) throws IOException {
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
      FormulaEnvironmentInput input, ExecutionInputBindings bindings) throws IOException {
    List<ExcelFormulaExternalWorkbookBinding> bindingsByName = new ArrayList<>();
    for (FormulaExternalWorkbookInput externalWorkbook : input.externalWorkbooks()) {
      bindingsByName.add(toExcelFormulaExternalWorkbookBinding(externalWorkbook, bindings));
    }
    return List.copyOf(bindingsByName);
  }

  private static ExcelFormulaExternalWorkbookBinding toExcelFormulaExternalWorkbookBinding(
      FormulaExternalWorkbookInput input, ExecutionInputBindings bindings) throws IOException {
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
      dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy policy) {
    if (policy == dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy.ERROR) {
      return ExcelFormulaMissingWorkbookPolicy.ERROR;
    }
    return ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE;
  }

  private static ExcelFormulaUdfToolpack toExcelUdfToolpack(FormulaUdfToolpackInput input) {
    return new ExcelFormulaUdfToolpack(
        input.name(),
        input.functions().stream().map(FormulaEnvironmentConverter::toExcelUdfFunction).toList());
  }

  private static ExcelFormulaUdfFunction toExcelUdfFunction(FormulaUdfFunctionInput input) {
    return new ExcelFormulaUdfFunction(
        input.name(),
        input.minimumArgumentCount(),
        input.maximumArgumentCount(),
        input.formulaTemplate());
  }
}
