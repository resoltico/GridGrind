package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.excel.ExcelFormulaEnvironment;
import dev.erst.gridgrind.excel.ExcelFormulaExternalWorkbookBinding;
import dev.erst.gridgrind.excel.ExcelFormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.excel.ExcelFormulaUdfFunction;
import dev.erst.gridgrind.excel.ExcelFormulaUdfToolpack;
import java.nio.file.Path;

/** Converts request-scoped formula environment payloads into engine-owned evaluator models. */
final class FormulaEnvironmentConverter {
  private FormulaEnvironmentConverter() {}

  static ExcelFormulaEnvironment toExcelFormulaEnvironment(FormulaEnvironmentInput input) {
    if (input == null) {
      return ExcelFormulaEnvironment.defaults();
    }
    return new ExcelFormulaEnvironment(
        input.externalWorkbooks().stream()
            .map(FormulaEnvironmentConverter::toExcelFormulaExternalWorkbookBinding)
            .toList(),
        toExcelMissingWorkbookPolicy(input.missingWorkbookPolicy()),
        input.udfToolpacks().stream()
            .map(FormulaEnvironmentConverter::toExcelUdfToolpack)
            .toList());
  }

  private static ExcelFormulaExternalWorkbookBinding toExcelFormulaExternalWorkbookBinding(
      FormulaExternalWorkbookInput input) {
    return new ExcelFormulaExternalWorkbookBinding(input.workbookName(), Path.of(input.path()));
  }

  private static ExcelFormulaMissingWorkbookPolicy toExcelMissingWorkbookPolicy(
      dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy policy) {
    return switch (policy) {
      case ERROR -> ExcelFormulaMissingWorkbookPolicy.ERROR;
      case USE_CACHED_VALUE -> ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE;
    };
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
