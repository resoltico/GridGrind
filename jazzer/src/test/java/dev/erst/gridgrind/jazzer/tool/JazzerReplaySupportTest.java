package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
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
  void replayClassifiesNamedRangeShiftArtifactAsSuccess() {
    byte[] input = artifactBytes("named-range-shift-overwrite-invalid.b64");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.xlsxRoundTrip(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new XlsxRoundTripDetails(
            83,
            4,
            Map.of(
                "CREATE_SHEET", 1L,
                "FORCE_FORMULA_RECALCULATION_ON_OPEN", 2L,
                "SET_NAMED_RANGE", 1L),
            Map.of()),
        success.details());
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
            6,
            Map.of(
                "CLEAR_SHEET_PROTECTION", 1L,
                "CREATE_SHEET", 1L,
                "UNGROUP_COLUMNS", 4L),
            Map.of()),
        success.details());
  }

  @Test
  void replayClassifiesAppendRowFailureWorkflowArtifactUsingCurrentReplayContract() {
    byte[] input =
        fileBytes(
            "src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/OperationWorkflowFuzzTestInputs/executeWorkflow/append_row_failure.bin");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.protocolWorkflow(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new ProtocolWorkflowDetails(
            67, "NEW", "NONE", 0, Map.of(), Map.of(), 0, Map.of(), "SUCCESS"),
        success.details());
  }

  private static byte[] artifactBytes(String resourceName) {
    return Base64.getDecoder()
        .decode(
            new String(
                    resourceBytes("/dev/erst/gridgrind/jazzer/tool/" + resourceName),
                    StandardCharsets.US_ASCII)
                .replaceAll("\\s+", ""));
  }

  private static byte[] resourceBytes(String resourcePath) {
    try (InputStream inputStream =
        JazzerReplaySupportTest.class.getResourceAsStream(resourcePath)) {
      if (inputStream == null) {
        throw new IllegalStateException("missing replay artifact resource: " + resourcePath);
      }
      return inputStream.readAllBytes();
    } catch (IOException exception) {
      throw new UncheckedIOException(
          "failed to load replay artifact resource: " + resourcePath, exception);
    }
  }

  private static byte[] fileBytes(String relativePath) {
    try {
      return Files.readAllBytes(Path.of(relativePath));
    } catch (IOException exception) {
      throw new UncheckedIOException(
          "failed to load replay artifact file: " + relativePath, exception);
    }
  }
}
