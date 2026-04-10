package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.Map;
import org.junit.jupiter.api.Test;

/** Tests for replay-time classification of raw protocol-request fuzz inputs. */
class JazzerReplaySupportTest {
  @Test
  void expectationForCapturesStableOutcomeKindAndReplayDetails() {
    ReplayOutcome outcome =
        JazzerReplaySupport.replay(
            JazzerHarness.protocolRequest(), "{".getBytes(StandardCharsets.UTF_8));

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
            JazzerHarness.protocolRequest(), "{".getBytes(StandardCharsets.UTF_8));

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
            JazzerHarness.protocolRequest(),
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
            JazzerHarness.protocolRequest(),
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

  @Test
  void replayParsesWorkbookHealthWorkflowExample() {
    byte[] input =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "operations": [
            { "type": "ENSURE_SHEET", "sheetName": "Budget Review" },
            { "type": "ENSURE_SHEET", "sheetName": "Summary" },
            {
              "type": "SET_CELL",
              "sheetName": "Summary",
              "address": "A1",
              "value": { "type": "FORMULA", "formula": "'Budget Review'!B1" }
            }
          ],
          "reads": [
            { "type": "ANALYZE_WORKBOOK_FINDINGS", "requestId": "lint" },
            {
              "type": "GET_CELLS",
              "requestId": "summary-cells",
              "sheetName": "Summary",
              "addresses": ["A1"]
            }
          ]
        }
        """
            .getBytes(StandardCharsets.UTF_8);
    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.protocolRequest(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new ProtocolRequestDetails(
            input.length,
            "PARSED",
            "NEW",
            "NONE",
            3,
            Map.of("ENSURE_SHEET", 2L, "SET_CELL", 1L),
            Map.of(),
            2,
            Map.of("ANALYZE_WORKBOOK_FINDINGS", 1L, "GET_CELLS", 1L)),
        success.details());
  }

  @Test
  void replayClassifiesNamedRangeShiftArtifactAsExpectedInvalidRoundTripInput() {
    byte[] input = artifactBytes("named-range-shift-overwrite-invalid.b64");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.xlsxRoundTrip(), input);

    assertInstanceOf(ReplayOutcome.ExpectedInvalid.class, outcome);
    ReplayOutcome.ExpectedInvalid expectedInvalid = (ReplayOutcome.ExpectedInvalid) outcome;
    assertEquals("IllegalArgumentException", expectedInvalid.invalidKind());
    assertEquals(
        "SHIFT_ROWS cannot move named range 'Name9' on sheet 'I'; row structural edits that would overwrite or partially move range-backed named ranges are not supported",
        expectedInvalid.message());
    assertEquals(
        new XlsxRoundTripDetails(
            83,
            9,
            Map.of(
                "CREATE_SHEET", 4L,
                "FORCE_FORMULA_RECALCULATION_ON_OPEN", 2L,
                "RENAME_SHEET", 1L,
                "SET_NAMED_RANGE", 1L,
                "SHIFT_ROWS", 1L),
            Map.of()),
        expectedInvalid.details());
  }

  @Test
  void replayClassifiesPartialCollapsedColumnUngroupArtifactAsSuccess() {
    byte[] input = artifactBytes("partial-collapsed-column-ungroup-roundtrip-success.b64");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.xlsxRoundTrip(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new XlsxRoundTripDetails(
            94,
            7,
            Map.of(
                "CLEAR_SHEET_PROTECTION", 1L,
                "CREATE_SHEET", 1L,
                "GROUP_COLUMNS", 1L,
                "UNGROUP_COLUMNS", 4L),
            Map.of()),
        success.details());
  }

  private static byte[] artifactBytes(String resourceName) {
    try (InputStream inputStream =
        JazzerReplaySupportTest.class.getResourceAsStream(
            "/dev/erst/gridgrind/jazzer/tool/" + resourceName)) {
      if (inputStream == null) {
        throw new IllegalStateException("missing replay artifact resource: " + resourceName);
      }
      String base64 =
          new String(inputStream.readAllBytes(), StandardCharsets.US_ASCII).replaceAll("\\s+", "");
      return Base64.getDecoder().decode(base64);
    } catch (IOException exception) {
      throw new UncheckedIOException(
          "failed to load replay artifact resource: " + resourceName, exception);
    }
  }
}
