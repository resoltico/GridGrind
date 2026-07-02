package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.exc.InvalidTypeIdException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;

/** Focused JSON problem-surface and limit-behavior tests for {@link GridGrindJson}. */
@SuppressWarnings("NotJavadoc")
class GridGrindJsonProblemSurfaceTest {
  @Test
  void validatesNullArgumentsAndDoesNotCloseOutputStreams() throws IOException {
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindJson.readRequest((InputStream) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindJson.readResponse((byte[]) null))
            .getMessage());
    assertEquals(
        "catalog must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJsonOutput.writeProtocolCatalogBytes(null))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJsonOutput.writeTypeEntry(new ByteArrayOutputStream(), null))
            .getMessage());

    try (TrackingOutputStream outputStream = new TrackingOutputStream()) {
      GridGrindJsonOutput.writeRequest(
          outputStream,
          WorkbookPlan.standard(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.None(),
              dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
              dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
              List.of()),
          false);
      assertFalse(outputStream.closed);
    }
  }

  @Test
  void rejectsOversizedRequestPayloads() {
    byte[] oversized =
        ("{\"source\":{\"type\":\"NEW\"},\"steps\":[],\"pad\":\""
                + "x".repeat((int) GridGrindJson.maxRequestDocumentBytes())
                + "\"}")
            .getBytes(StandardCharsets.UTF_8);

    InvalidRequestException failure =
        assertThrows(InvalidRequestException.class, () -> GridGrindJson.readRequest(oversized));

    assertEquals(
        "Request JSON exceeds the maximum size of 16 MiB (16777216 bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT sources.",
        failure.getMessage());
  }

  @Test
  void rejectsOversizedRequestStreamsWithTheSameProductOwnedMessage() {
    byte[] oversized =
        ("{\"planId\":\""
                + "x".repeat((int) GridGrindJson.maxRequestDocumentBytes())
                + "\",\"source\":{\"type\":\"NEW\"},\"steps\":[]}")
            .getBytes(StandardCharsets.UTF_8);

    InvalidRequestException failure =
        assertThrows(
            InvalidRequestException.class,
            () -> GridGrindJson.readRequest(new ByteArrayInputStream(oversized)));

    assertEquals(
        "Request JSON exceeds the maximum size of 16 MiB (16777216 bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT sources.",
        failure.getMessage());
  }

  @Test
  void readsInvalidResponsesAndCatalogsUsingPublicProblemTypes() {
    InvalidJsonException invalidResponse =
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readResponse("{".getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException invalidCatalog =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readProtocolCatalog(
                    """
                    {
                      "protocolVersion": "V1",
                      "requestTemplate": { "source": { "type": "NEW" }, "steps": [] }
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals(Optional.of(1), invalidResponse.jsonLine());
    assertEquals(Optional.of(2), invalidResponse.jsonColumn());
    assertTrue(invalidCatalog.getMessage().startsWith("Missing required field '"));
  }

  @Test
  void exposesProductOwnedHelperMessagesAndJsonLocations() throws IOException {
    JsonFactory jsonFactory = new JsonFactory();
    MismatchedInputException floatingInteger =
        (MismatchedInputException)
            MismatchedInputException.from(
                    utf8Parser(jsonFactory, "2.5"),
                    Integer.class,
                    "Cannot coerce Floating-point value (2.5) to `int` value"
                        + " (but could if coercion was enabled using `CoercionConfig`)")
                .prependPath(WorkbookPlan.class, "rowCount");
    InvalidTypeIdException invalidType =
        InvalidTypeIdException.from(utf8Parser(jsonFactory, "\"x\""), "bad type", null, "NOPE");

    assertEquals(
        "Field 'rowCount' must be an integer value",
        GridGrindJson.mismatchedInputMessage(floatingInteger));
    assertEquals("Unknown type value 'NOPE'", GridGrindJson.message(invalidType));
    assertEquals(
        "Unknown field 'reads'",
        GridGrindJson.message(
            UnrecognizedPropertyException.from(
                utf8Parser(jsonFactory, "{}"), WorkbookPlan.class, "reads", List.of())));
    assertEquals(
        "Cannot construct instance of `x`",
        GridGrindJson.message(new IllegalArgumentException("Cannot construct instance of `x`")));
    assertEquals(
        "Cannot deserialize value as a subtype of `x` (for POJO property 'target')"
            + " (but could if coercion was enabled using `CoercionConfig`)",
        GridGrindJson.cleanJacksonMessage(
            "Cannot deserialize value as a subtype of `x` (for POJO property 'target')"
                + " (but could if coercion was enabled using `CoercionConfig`)"));
    assertEquals("Invalid JSON payload", GridGrindJson.cleanJacksonMessage(" "));
    assertEquals(Optional.empty(), GridGrindJson.jsonLine(null));
    assertEquals(Optional.empty(), GridGrindJson.jsonColumn(null));
    assertEquals(
        Optional.of(4), GridGrindJson.jsonLine(new TokenStreamLocation(null, 0L, 0L, 4, 9)));
    assertEquals(
        Optional.of(9), GridGrindJson.jsonColumn(new TokenStreamLocation(null, 0L, 0L, 4, 9)));
  }

  @Test
  void surfacesSimilarValidTypeIdsForTyposInKnownActionTypes() {
    InvalidRequestShapeException typo =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "typo",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "action": { "type": "COVE_SHEET" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(
        typo.getMessage().contains("MOVE_SHEET") && typo.getMessage().contains("COPY_SHEET"),
        "typo matching multiple action types should list all similar candidates");
  }

  @Test
  void surfacesUnknownTypeMessageForAtJsonSubTypesAnnotatedFields() {
    InvalidRequestShapeException badAnchor =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "bad-anchor",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "action": {
                            "type": "SET_CHART",
                            "chart": {
                              "name": "C",
                              "anchor": { "type": "ONE_CELL" },
                              "title": { "type": "NONE" },
                              "legend": { "type": "VISIBLE", "position": "RIGHT" },
                              "displayBlanksAs": "GAP",
                              "plotOnlyVisibleCells": true,
                              "plots": [{"type": "BAR", "series": []}]
                            }
                          }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(
        badAnchor.getMessage().startsWith("Unknown type value 'ONE_CELL'"),
        "unknown anchor type from @JsonSubTypes-annotated field should be reported");
  }

  @Test
  void requestUnknownFieldDiagnosticsPointAtTheOffendingFieldOnce() {
    InvalidRequestShapeException topLevelUnknownField =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": [],
                      "bogus": 1
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException nestedUnknownField =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": [
                        {
                          "stepId": "summary",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "query": {
                            "type": "GET_WORKBOOK_SUMMARY",
                            "extra": true
                          }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Unknown field 'bogus'", topLevelUnknownField.getMessage());
    assertEquals(Optional.of("bogus"), topLevelUnknownField.jsonPath());
    assertEquals("Unknown field 'steps[0].query.extra'", nestedUnknownField.getMessage());
    assertEquals(Optional.of("steps[0].query.extra"), nestedUnknownField.jsonPath());
  }

  /** Tracks whether the request writer closes the destination stream after producing JSON. */
  private static final class TrackingOutputStream extends ByteArrayOutputStream {
    private boolean closed;

    @Override
    public void close() {
      closed = true;
    }
  }

  private static tools.jackson.core.JsonParser utf8Parser(JsonFactory jsonFactory, String json)
      throws IOException {
    return jsonFactory.createParser(
        tools.jackson.core.ObjectReadContext.empty(),
        new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
  }
}
