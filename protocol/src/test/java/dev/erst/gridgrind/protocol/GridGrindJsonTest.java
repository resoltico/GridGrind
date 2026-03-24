package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for GridGrindJson serialization and deserialization. */
class GridGrindJsonTest {
  @Test
  void defaultsRequestProtocolVersionDuringJsonRead() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "mode": "NEW" },
              "operations": [],
              "analysis": { "sheets": [] }
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(GridGrindProtocolVersion.V1, request.protocolVersion());
  }

  @Test
  void readsRequestsFromInputStreamsWithoutClosingThem() throws IOException {
    try (TrackingInputStream inputStream =
        new TrackingInputStream(
            """
            {
              "protocolVersion": "V1",
              "source": { "mode": "NEW" },
              "operations": [],
              "analysis": { "sheets": [] }
            }
            """
                .getBytes(StandardCharsets.UTF_8))) {
      GridGrindRequest request = GridGrindJson.readRequest(inputStream);

      assertEquals(GridGrindProtocolVersion.V1, request.protocolVersion());
      assertFalse(inputStream.closed);
    }
  }

  @Test
  void roundTripsStructuredResponses() throws IOException {
    GridGrindResponse expected =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            "/tmp/budget.xlsx",
            new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), true),
            List.of(new GridGrindResponse.SheetReport("Budget", 1, 0, 1, List.of(), List.of())));

    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    GridGrindJson.writeResponse(outputStream, expected);

    GridGrindResponse streamDecoded =
        GridGrindJson.readResponse(new ByteArrayInputStream(outputStream.toByteArray()));
    GridGrindResponse actual = GridGrindJson.readResponse(outputStream.toByteArray());
    byte[] encoded = GridGrindJson.writeResponseBytes(expected);

    assertEquals(expected, streamDecoded);
    assertEquals(expected, actual);
    assertArrayEquals(encoded, outputStream.toByteArray());
  }

  @Test
  void wrapsMalformedJsonAsInvalidPayloadErrors() {
    InvalidJsonException requestFailure =
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readRequest("{".getBytes(StandardCharsets.UTF_8)));
    InvalidJsonException requestStreamFailure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readRequest(
                    new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8))));
    InvalidJsonException responseFailure =
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readResponse("{".getBytes(StandardCharsets.UTF_8)));
    InvalidJsonException responseStreamFailure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readResponse(
                    new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8))));

    assertNotNull(requestFailure.getMessage());
    assertNotNull(requestStreamFailure.getMessage());
    assertNotNull(responseFailure.getMessage());
    assertNotNull(responseStreamFailure.getMessage());
    assertFalse(requestFailure.getMessage().contains("[Source:"));
    assertFalse(requestFailure.getMessage().contains("reference chain"));
    assertNull(requestFailure.jsonPath());
    assertEquals(1, requestFailure.jsonLine());
    assertEquals(2, requestFailure.jsonColumn());
  }

  @Test
  void wrapsRequestValidationFailuresAsInvalidRequestErrors() {
    InvalidRequestException requestFailure =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                {
                  "source": { "mode": "NEW" },
                  "operations": [],
                  "analysis": {
                    "sheets": [
                      { "sheetName": "Budget", "previewRowCount": 0, "previewColumnCount": 1 }
                    ]
                  }
                }
                """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("previewRowCount must be greater than 0", requestFailure.getMessage());
    assertEquals("analysis.sheets[0]", requestFailure.jsonPath());
    assertEquals(6, requestFailure.jsonLine());
    assertEquals(78, requestFailure.jsonColumn());
  }

  @Test
  void validatesNullArguments() {
    assertThrows(NullPointerException.class, () -> GridGrindJson.readRequest((byte[]) null));
    assertThrows(
        NullPointerException.class, () -> GridGrindJson.readRequest((ByteArrayInputStream) null));
    assertThrows(NullPointerException.class, () -> GridGrindJson.readResponse((byte[]) null));
    assertThrows(
        NullPointerException.class, () -> GridGrindJson.readResponse((ByteArrayInputStream) null));
    assertThrows(NullPointerException.class, () -> GridGrindJson.writeResponseBytes(null));
    assertThrows(
        NullPointerException.class,
        () ->
            GridGrindJson.writeResponse(
                null,
                new GridGrindResponse.Failure(
                    null,
                    GridGrindResponse.Problem.of(
                        GridGrindProblemCode.INVALID_REQUEST,
                        "bad",
                        new GridGrindResponse.ProblemContext.ValidateRequest(null, null)))));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeResponse(new ByteArrayOutputStream(), null));
  }

  private static final class TrackingInputStream extends ByteArrayInputStream {
    private boolean closed;

    private TrackingInputStream(byte[] bytes) {
      super(bytes);
    }

    @Override
    public void close() throws IOException {
      closed = true;
      super.close();
    }
  }
}
