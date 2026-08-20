package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
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
}
