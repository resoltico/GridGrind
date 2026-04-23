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
    return toExcelFormulaEnvironment(input, Path.of(""));
  }

  static ExcelFormulaEnvironment toExcelFormulaEnvironment(
      FormulaEnvironmentInput input, Path workingDirectory) {
    if (input == null) {
      return ExcelFormulaEnvironment.defaults();
    }
    return new ExcelFormulaEnvironment(
        input.externalWorkbooks().stream()
            .map(binding -> toExcelFormulaExternalWorkbookBinding(binding, workingDirectory))
            .toList(),
        toExcelMissingWorkbookPolicy(input.missingWorkbookPolicy()),
        input.udfToolpacks().stream()
            .map(FormulaEnvironmentConverter::toExcelUdfToolpack)
            .toList());
  }

  private static ExcelFormulaExternalWorkbookBinding toExcelFormulaExternalWorkbookBinding(
      FormulaExternalWorkbookInput input, Path workingDirectory) {
    return new ExcelFormulaExternalWorkbookBinding(
        input.workbookName(), ExecutionRequestPaths.normalizePath(input.path(), workingDirectory));
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
