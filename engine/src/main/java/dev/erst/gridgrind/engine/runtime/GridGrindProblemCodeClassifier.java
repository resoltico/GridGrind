package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.json.InvalidEncodingException;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.InvalidSigningConfigurationException;
import dev.erst.gridgrind.excel.InvalidWorkbookPasswordException;
import dev.erst.gridgrind.excel.MissingExternalWorkbookException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnregisteredUserDefinedFunctionException;
import dev.erst.gridgrind.excel.UnsupportedFormulaConstructException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookPasswordRequiredException;
import dev.erst.gridgrind.excel.WorkbookSecurityException;
import java.io.IOException;
import java.time.DateTimeException;
import java.util.Objects;

/** Maps runtime and protocol exceptions onto public GridGrind problem codes. */
final class GridGrindProblemCodeClassifier {
  private GridGrindProblemCodeClassifier() {}

  static GridGrindProblemCode codeFor(Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    return switch (exception) {
      case InvalidEncodingException _ -> GridGrindProblemCode.INVALID_ENCODING;
      case InvalidJsonException _ -> GridGrindProblemCode.INVALID_JSON;
      case InvalidRequestShapeException _ -> GridGrindProblemCode.INVALID_REQUEST_SHAPE;
      case InvalidRequestException _ -> GridGrindProblemCode.INVALID_REQUEST;
      case RequestPathEscapeException _ -> GridGrindProblemCode.PATH_ESCAPES_ROOT;
      case UnsafePathAccessException _ -> GridGrindProblemCode.UNSAFE_PATH_ACCESS;
      case AssertionFailedException _ -> GridGrindProblemCode.ASSERTION_FAILED;
      case WorkbookNotFoundException _ -> GridGrindProblemCode.WORKBOOK_NOT_FOUND;
      case InputSourceNotFoundException _ -> GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND;
      case InputSourceUnavailableException _ -> GridGrindProblemCode.INPUT_SOURCE_UNAVAILABLE;
      case InputSourceReadException _ -> GridGrindProblemCode.INPUT_SOURCE_IO_ERROR;
      case SheetNotFoundException _ -> GridGrindProblemCode.SHEET_NOT_FOUND;
      case NamedRangeNotFoundException _ -> GridGrindProblemCode.NAMED_RANGE_NOT_FOUND;
      case CellNotFoundException _ -> GridGrindProblemCode.CELL_NOT_FOUND;
      case InvalidCellAddressException _ -> GridGrindProblemCode.INVALID_CELL_ADDRESS;
      case InvalidRangeAddressException _ -> GridGrindProblemCode.INVALID_RANGE_ADDRESS;
      case InvalidFormulaException _ -> GridGrindProblemCode.INVALID_FORMULA;
      case UnsupportedFormulaConstructException _ ->
          GridGrindProblemCode.UNSUPPORTED_FORMULA_CONSTRUCT;
      case MissingExternalWorkbookException _ -> GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK;
      case UnregisteredUserDefinedFunctionException _ ->
          GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION;
      case UnsupportedFormulaException _ -> GridGrindProblemCode.UNSUPPORTED_FORMULA;
      case WorkbookPasswordRequiredException _ -> GridGrindProblemCode.WORKBOOK_PASSWORD_REQUIRED;
      case InvalidWorkbookPasswordException _ -> GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD;
      case InvalidSigningConfigurationException _ ->
          GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION;
      case WorkbookSecurityException _ -> GridGrindProblemCode.WORKBOOK_SECURITY_ERROR;
      case IOException _ -> GridGrindProblemCode.IO_ERROR;
      case IllegalArgumentException _ -> GridGrindProblemCode.INVALID_REQUEST;
      case DateTimeException _ -> GridGrindProblemCode.INVALID_REQUEST;
      default -> GridGrindProblemCode.INTERNAL_ERROR;
    };
  }
}
