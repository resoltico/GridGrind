package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.problem.InvalidEncodingException;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.RequestDuplicateKey;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.RequestInvalidJson;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Verifies typed request conversion uses the canonical tolerant structural analysis. */
class GridGrindRequestDecoderTest {
  @Test
  void routesPresenceAndDuplicateFailuresBeforeTypedBindingAndOtherShapeFailuresThroughIt() {
    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWith("\"planId\": \"one\", \"planId\": \"two\""))));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {"protocolVersion":"V3","source":{"type":"NEW"},"persistence":{"type":"NONE"}}
                    """)));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {"protocolVersion":"V3","source":{"type":"NEW"},"persistence":{"type":"NONE"},"steps":null}
                    """)));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {"protocolVersion":"V3","source":{},"persistence":{"type":"NONE"},"steps":[]}
                    """)));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () -> GridGrindJson.readRequest(requestWith("\"unknown\": true"))));
    assertInstanceOf(
        InvalidRequestShapeException.class,
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {"protocolVersion":"V3","source":{"type":7},"persistence":{"type":"NONE"},"steps":[]}
                    """)));
  }

  @Test
  void preservesTheDedicatedInvalidEncodingExceptionBeforeJsonTreeParsing() {
    InvalidEncodingException failure =
        assertThrows(
            InvalidEncodingException.class,
            () -> GridGrindJson.readRequest(new byte[] {'{', (byte) 0xc3, (byte) 0x28}));

    assertEquals("Request bytes must be valid UTF-8", failure.getMessage());
  }

  @Test
  void mapsBothInvalidJsonSourcesThroughTheSameCanonicalExceptionFamily() {
    assertInstanceOf(
        InvalidJsonException.class,
        new RequestInvalidJson(
                "invalid JSON", java.util.Optional.empty(), java.util.Optional.of(0L))
            .exception());
    assertInstanceOf(
        InvalidJsonException.class, new RequestDuplicateKey("", "planId", 0, 0).exception());
  }

  @Test
  void selectsTheFirstCanonicalTolerantProblemWithoutStrictlyReparsingTheRequest() {
    byte[] request =
        """
        {
          "protocolVersion": "V3",
          "source" { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8);
    RequestIntake analysis = GridGrindJson.analyzeRequest(request);

    InvalidJsonException failure =
        assertThrows(InvalidJsonException.class, () -> GridGrindJson.readRequest(request));

    assertEquals(analysis.structuralProblems().getFirst().message(), failure.getMessage());
  }

  private static byte[] requestWith(String replacement) {
    return """
        {
          "protocolVersion": "V3",
          "source": {"type": "NEW"},
          "persistence": {"type": "NONE"},
          %s,
          "steps": []
        }
        """
        .formatted(replacement)
        .getBytes(StandardCharsets.UTF_8);
  }
}
