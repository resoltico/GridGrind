package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.InvalidFormulaInputException;
import dev.erst.gridgrind.contract.dto.InvalidRawFormulaTextException;
import dev.erst.gridgrind.contract.json.FormulaRequestException;
import dev.erst.gridgrind.contract.json.InvalidEncodingException;
import dev.erst.gridgrind.contract.json.NumberNotRepresentableException;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Guards classification of request-byte decoding failures at the engine boundary. */
class GridGrindProblemCodeClassifierTest {
  @Test
  void mapsInvalidUtf8ToTheDedicatedProtocolProblemCode() {
    assertEquals(
        GridGrindProblemCode.INVALID_ENCODING,
        GridGrindProblemCodeClassifier.codeFor(
            new InvalidEncodingException(
                "Request bytes must be valid UTF-8",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                null)));
  }

  @Test
  void mapsRequestPathFailuresToTheirDedicatedProtocolCodes() {
    assertEquals(
        GridGrindProblemCode.PATH_ESCAPES_ROOT,
        GridGrindProblemCodeClassifier.codeFor(
            new RequestPathEscapeException("path escapes root")));
    assertEquals(
        GridGrindProblemCode.UNSAFE_PATH_ACCESS,
        GridGrindProblemCodeClassifier.codeFor(new UnsafePathAccessException("no safe binding")));
  }

  @Test
  void mapsLossyNumericTokensToTheDedicatedProtocolCode() {
    assertEquals(
        GridGrindProblemCode.NUMBER_NOT_REPRESENTABLE,
        GridGrindProblemCodeClassifier.codeFor(
            new NumberNotRepresentableException("Use TEXT", "steps[0].action.value.number")));
  }

  @Test
  void mapsBothFormulaInputContractsToTheirOwnedProblemCodes() {
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        GridGrindProblemCodeClassifier.codeFor(
            new FormulaRequestException(
                GridGrindProblemCode.INVALID_FORMULA,
                "formula must not begin with =",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                null)));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA_TEXT,
        GridGrindProblemCodeClassifier.codeFor(
            new FormulaRequestException(
                GridGrindProblemCode.INVALID_FORMULA_TEXT,
                "formula text contains a forbidden XML character",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                null)));
    assertEquals(
        "problemCode must classify formula input",
        org.junit.jupiter.api.Assertions.assertThrows(
                IllegalArgumentException.class,
                () ->
                    new FormulaRequestException(
                        GridGrindProblemCode.IO_ERROR,
                        "not a formula problem",
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        null))
            .getMessage());
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        GridGrindProblemCodeClassifier.codeFor(
            new InvalidFormulaInputException("formula must not begin with =")));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA_TEXT,
        GridGrindProblemCodeClassifier.codeFor(
            new InvalidRawFormulaTextException("formula contains a forbidden XML character")));
  }
}
