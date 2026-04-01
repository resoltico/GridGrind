package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.nio.charset.StandardCharsets;
import java.util.Map;
import org.junit.jupiter.api.Test;

/** Tests for replay-time classification of raw protocol-request fuzz inputs. */
class JazzerReplaySupportTest {
  @Test
  void expectationForCapturesStableOutcomeKindAndReplayDetails() {
    ReplayOutcome outcome =
        JazzerReplaySupport.replay(
            JazzerHarness.PROTOCOL_REQUEST, "{".getBytes(StandardCharsets.UTF_8));

    assertEquals(
        new ReplayExpectation(
            "EXPECTED_INVALID",
            new ProtocolRequestDetails(
                1, "INVALID_JSON", "NOT_PARSED", "NOT_PARSED", 0, Map.of(), Map.of(), 0, Map.of())),
        JazzerReplaySupport.expectationFor(outcome));
  }

  @Test
  void replayClassifiesMalformedJsonAsInvalidJson() {
    ReplayOutcome outcome =
        JazzerReplaySupport.replay(
            JazzerHarness.PROTOCOL_REQUEST, "{".getBytes(StandardCharsets.UTF_8));

    assertInstanceOf(ReplayOutcome.ExpectedInvalid.class, outcome);
    ReplayOutcome.ExpectedInvalid expectedInvalid = (ReplayOutcome.ExpectedInvalid) outcome;
    assertEquals("InvalidJsonException", expectedInvalid.invalidKind());
    assertEquals(
        "INVALID_JSON", ((ProtocolRequestDetails) expectedInvalid.details()).decodeOutcome());
  }

  @Test
  void replayClassifiesRequestShapeFailuresSeparately() {
    ReplayOutcome outcome =
        JazzerReplaySupport.replay(
            JazzerHarness.PROTOCOL_REQUEST,
            """
            {
              "protocolVersion": "V1",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "operations": [],
              "reads": [
                { "type": "GET_WORKBOOK_SUMMARY" }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertInstanceOf(ReplayOutcome.ExpectedInvalid.class, outcome);
    ReplayOutcome.ExpectedInvalid expectedInvalid = (ReplayOutcome.ExpectedInvalid) outcome;
    assertEquals("InvalidRequestShapeException", expectedInvalid.invalidKind());
    assertEquals(
        "INVALID_REQUEST_SHAPE",
        ((ProtocolRequestDetails) expectedInvalid.details()).decodeOutcome());
  }

  @Test
  void replayKeepsSemanticRequestFailuresAsInvalidRequest() {
    ReplayOutcome outcome =
        JazzerReplaySupport.replay(
            JazzerHarness.PROTOCOL_REQUEST,
            """
            {
              "protocolVersion": "V1",
              "source": { "type": "EXISTING", "path": "budget.xlsm" },
              "persistence": { "type": "NONE" },
              "operations": [],
              "reads": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertInstanceOf(ReplayOutcome.ExpectedInvalid.class, outcome);
    ReplayOutcome.ExpectedInvalid expectedInvalid = (ReplayOutcome.ExpectedInvalid) outcome;
    assertEquals("InvalidRequestException", expectedInvalid.invalidKind());
    assertEquals(
        "INVALID_REQUEST", ((ProtocolRequestDetails) expectedInvalid.details()).decodeOutcome());
  }
}
