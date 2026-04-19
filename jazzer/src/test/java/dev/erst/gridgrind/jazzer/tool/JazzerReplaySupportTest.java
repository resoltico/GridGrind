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
                1,
                "INVALID_JSON",
                "NOT_PARSED",
                "NOT_PARSED",
                0,
                Map.of(),
                Map.of(),
                0,
                Map.of(),
                0,
                Map.of())),
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
              "steps": [
                {
                  "query": { "type": "GET_WORKBOOK_SUMMARY" },
                  "target": { "type": "CURRENT" }
                }
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
              "steps": []
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
    byte[] input = fileBytes("examples/workbook-health-request.json");
    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.protocolRequest(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new ProtocolRequestDetails(
            input.length,
            "PARSED",
            "NEW",
            "NONE",
            5,
            Map.of("ENSURE_SHEET", 2L, "SET_CELL", 3L),
            Map.of(),
            0,
            Map.of(),
            4,
            Map.of(
                "ANALYZE_FORMULA_HEALTH", 1L,
                "ANALYZE_WORKBOOK_FINDINGS", 1L,
                "GET_CELLS", 1L,
                "GET_SHEET_SUMMARY", 1L)),
        success.details());
  }

  @Test
  void replayParsesLargeFileModesExample() {
    byte[] input = fileBytes("examples/large-file-modes-request.json");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.protocolRequest(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new ProtocolRequestDetails(
            input.length,
            "PARSED",
            "NEW",
            "SAVE_AS",
            4,
            Map.of("APPEND_ROW", 3L, "ENSURE_SHEET", 1L),
            Map.of(),
            0,
            Map.of(),
            2,
            Map.of("GET_SHEET_SUMMARY", 1L, "GET_WORKBOOK_SUMMARY", 1L)),
        success.details());
  }

  @Test
  void replayParsesPackageSecuritySeed() {
    byte[] input =
        fileBytes(
            "src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/ProtocolRequestFuzzTestInputs/readRequest/package_security_request.json");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.protocolRequest(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new ProtocolRequestDetails(
            input.length,
            "PARSED",
            "EXISTING",
            "SAVE_AS",
            1,
            Map.of("SET_CELL", 1L),
            Map.of(),
            0,
            Map.of(),
            2,
            Map.of("GET_CELLS", 1L, "GET_PACKAGE_SECURITY", 1L)),
        success.details());
  }

  @Test
  void replayClassifiesNamedRangeShiftArtifactAsExpectedInvalid() {
    byte[] input = artifactBytes("named-range-shift-overwrite-invalid.b64");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.xlsxRoundTrip(), input);

    assertInstanceOf(ReplayOutcome.ExpectedInvalid.class, outcome);
    ReplayOutcome.ExpectedInvalid expectedInvalid = (ReplayOutcome.ExpectedInvalid) outcome;
    assertEquals("SheetNotFoundException", expectedInvalid.invalidKind());
    assertEquals("Sheet does not exist: W", expectedInvalid.message());
    assertEquals(
        new XlsxRoundTripDetails(
            83,
            4,
            Map.of("AUTO_SIZE_COLUMNS", 2L, "CREATE_SHEET", 1L, "SET_NAMED_RANGE", 1L),
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
            6,
            Map.of(
                "CLEAR_SHEET_PROTECTION", 1L,
                "CREATE_SHEET", 1L,
                "UNGROUP_COLUMNS", 4L),
            Map.of()),
        success.details());
  }

  @Test
  void replayClassifiesProtocolWorkflowCase01UsingCurrentReplayContract() {
    byte[] input =
        fileBytes(
            "src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/OperationWorkflowFuzzTestInputs/executeWorkflow/workflow_case_01.bin");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.protocolWorkflow(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new ProtocolWorkflowDetails(
            67, "NEW", "NONE", 0, Map.of(), Map.of(), 0, Map.of(), 0, Map.of(), "SUCCESS"),
        success.details());
  }

  @Test
  void replayClassifiesProtocolWorkflowCase09UsingCurrentReplayContract() {
    byte[] input =
        fileBytes(
            "src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/OperationWorkflowFuzzTestInputs/executeWorkflow/workflow_case_09.bin");

    ReplayOutcome outcome = JazzerReplaySupport.replay(JazzerHarness.protocolWorkflow(), input);

    assertInstanceOf(ReplayOutcome.Success.class, outcome);
    ReplayOutcome.Success success = (ReplayOutcome.Success) outcome;
    assertEquals(
        new ProtocolWorkflowDetails(
            314,
            "EXISTING",
            "SAVE_AS",
            5,
            Map.of("APPLY_STYLE", 2L, "AUTO_SIZE_COLUMNS", 3L),
            Map.ofEntries(
                Map.entry("font", 2L),
                Map.entry("text_rotation", 2L),
                Map.entry("fill_patterned", 2L),
                Map.entry("border_top_color", 2L),
                Map.entry("border_bottom_color", 2L),
                Map.entry("horizontal_alignment", 2L),
                Map.entry("font_color", 2L),
                Map.entry("locked", 2L),
                Map.entry("protection", 2L),
                Map.entry("border_bottom", 2L),
                Map.entry("font_name", 2L),
                Map.entry("border_top", 2L),
                Map.entry("underline", 2L),
                Map.entry("strikeout", 2L),
                Map.entry("vertical_alignment", 2L),
                Map.entry("fill_pattern", 2L),
                Map.entry("indentation", 2L),
                Map.entry("alignment", 2L),
                Map.entry("italic", 1L),
                Map.entry("hidden_formula", 2L),
                Map.entry("fill", 2L),
                Map.entry("border", 2L),
                Map.entry("bold", 1L),
                Map.entry("wrap_text", 2L),
                Map.entry("font_height", 2L),
                Map.entry("font_height_twips", 2L)),
            3,
            Map.of("EXPECT_PRESENT", 3L),
            6,
            Map.of("ANALYZE_WORKBOOK_FINDINGS", 6L),
            "FAILURE"),
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
      return Files.readAllBytes(resolveProjectPath(relativePath));
    } catch (IOException exception) {
      throw new UncheckedIOException(
          "failed to load replay artifact file: " + relativePath, exception);
    }
  }

  private static Path resolveProjectPath(String relativePath) {
    Path moduleRelativePath = Path.of(relativePath);
    if (Files.exists(moduleRelativePath)) {
      return moduleRelativePath;
    }
    Path repoRelativePath = Path.of("..").resolve(relativePath).normalize();
    if (Files.exists(repoRelativePath)) {
      return repoRelativePath;
    }
    throw new IllegalStateException("missing replay artifact file: " + relativePath);
  }
}
