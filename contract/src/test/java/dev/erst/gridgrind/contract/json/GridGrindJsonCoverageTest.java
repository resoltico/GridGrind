package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.time.DateTimeException;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.exc.MismatchedInputException;

/** Additional parser-wording and invalid-payload coverage for the shared JSON codec. */
class GridGrindJsonCoverageTest {
  @Test
  void readsResponsesAndCatalogsFromStreamsWithoutClosingThem() throws IOException {
    GridGrindResponse response =
        new GridGrindResponse.Success(
            null,
            null,
            null,
            null,
            List.of(
                new dev.erst.gridgrind.contract.query.InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();

    try (TrackingInputStream responseStream =
            new TrackingInputStream(GridGrindJson.writeResponseBytes(response));
        TrackingInputStream catalogStream =
            new TrackingInputStream(GridGrindJson.writeProtocolCatalogBytes(catalog))) {
      assertEquals(response, GridGrindJson.readResponse(responseStream));
      assertEquals(catalog, GridGrindJson.readProtocolCatalog(catalogStream));
      assertFalse(responseStream.closed);
      assertFalse(catalogStream.closed);
    }
  }

  @Test
  void invalidResponseAndCatalogStreamsSurfaceInvalidJsonWithoutClosingCallerStreams()
      throws IOException {
    try (TrackingInputStream responseStream =
            new TrackingInputStream("{".getBytes(StandardCharsets.UTF_8));
        TrackingInputStream catalogStream =
            new TrackingInputStream("{".getBytes(StandardCharsets.UTF_8))) {
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class, () -> GridGrindJson.readResponse(responseStream)));
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class, () -> GridGrindJson.readProtocolCatalog(catalogStream)));
      assertFalse(responseStream.closed);
      assertFalse(catalogStream.closed);
    }
  }

  @Test
  void validatesNullArgumentsAcrossAllPublicReadAndWriteSurfaceMethods() {
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindJson.readResponse((InputStream) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.readProtocolCatalog((InputStream) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindJson.readProtocolCatalog((byte[]) null))
            .getMessage());
    assertEquals(
        "request must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindJson.writeRequestBytes(null))
            .getMessage());
    assertEquals(
        "response must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindJson.writeResponseBytes(null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJson.writeResponse(
                        null,
                        new GridGrindResponse.Failure(
                            null,
                            new GridGrindResponse.Problem(
                                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_JSON,
                                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_JSON
                                    .category(),
                                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_JSON
                                    .recovery(),
                                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_JSON
                                    .title(),
                                "bad",
                                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_JSON
                                    .resolution(),
                                new GridGrindResponse.ProblemContext.ParseArguments("--request"),
                                null,
                                List.of()))))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.writeProtocolCatalog(null, GridGrindProtocolCatalog.catalog()))
            .getMessage());
  }

  @Test
  void helperMethodsCoverFallbackAndLocationEdgeCases() throws IOException {
    JsonFactory jsonFactory = new JsonFactory();
    MismatchedInputException nullPrimitiveWithoutPath =
        MismatchedInputException.from(
            jsonFactory.createParser("null"), Integer.class, "Cannot map `null` into type `int`");
    MismatchedInputException nullPrimitiveWithIndex =
        (MismatchedInputException)
            MismatchedInputException.from(
                    jsonFactory.createParser("null"),
                    Integer.class,
                    "Cannot map `null` into type `int`")
                .prependPath(new Object(), 0);
    MismatchedInputException floatingPointWithoutPath =
        MismatchedInputException.from(
            jsonFactory.createParser("2.5"),
            Integer.class,
            "Floating-point value (2.5) out of range of int");
    MismatchedInputException floatingPointWithIndex =
        (MismatchedInputException)
            MismatchedInputException.from(
                    jsonFactory.createParser("2.5"),
                    Integer.class,
                    "Floating-point value (2.5) out of range of int")
                .prependPath(new Object(), 0);

    assertEquals(
        "Missing required field", GridGrindJson.mismatchedInputMessage(nullPrimitiveWithoutPath));
    assertEquals(
        "Missing required field", GridGrindJson.mismatchedInputMessage(nullPrimitiveWithIndex));
    assertEquals(
        "JSON value must be an integer value",
        GridGrindJson.mismatchedInputMessage(floatingPointWithoutPath));
    assertEquals(
        "JSON value must be an integer value",
        GridGrindJson.mismatchedInputMessage(floatingPointWithIndex));
    assertEquals(
        "Missing required field 'fieldName'",
        GridGrindJson.message(new IllegalArgumentException("problem: fieldName must not be null")));
    assertEquals(
        "JSON value has the wrong shape for this field",
        GridGrindJson.message(
            new IllegalArgumentException("Cannot deserialize value of type `x` from String")));
    assertEquals("bad", invokeInvalidPayload(new StreamConstraintsException("bad")).getMessage());
    assertEquals(
        "2026-04-17T25:00:00 is not a valid date",
        invokeInvalidPayload(
                new WrappedJacksonException(
                    "wrapper", new DateTimeException("2026-04-17T25:00:00 is not a valid date")))
            .getMessage());
    assertEquals(
        "Invalid JSON payload",
        GridGrindJson.cleanJacksonMessage(
            " (start marker at [Source: REDACTED; line: 1, column: 1])"));
    assertEquals("Invalid JSON payload", GridGrindJson.cleanJacksonMessage(null));
    assertEquals(
        "Missing required field 'fieldName'",
        GridGrindJson.message(
            new IllegalStateException("Missing required creator property 'fieldName'")));
    assertEquals(
        "Missing required field 'type'",
        GridGrindJson.message(new IllegalStateException("missing type id property 'type'")));
    assertEquals(
        "Invalid JSON payload",
        GridGrindJson.mismatchedInputMessage(
            MismatchedInputException.from(
                jsonFactory.createParser("null"), Integer.class, (String) null)));
    assertNull(GridGrindJson.jsonLine(new TokenStreamLocation(null, 0L, 0L, 0, 9)));
    assertNull(GridGrindJson.jsonColumn(new TokenStreamLocation(null, 0L, 0L, 4, 0)));
  }

  @Test
  void typeEntryAndRequestReadersStayDeterministicThroughPublicRoundTrip() throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    TypeEntry entry = GridGrindProtocolCatalog.entryFor("GET_CELLS").orElseThrow();
    GridGrindJson.writeTypeEntry(outputStream, entry);
    Catalog catalog =
        GridGrindJson.readProtocolCatalog(
            new ByteArrayInputStream(
                GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog())));
    WorkbookPlan template =
        GridGrindJson.readRequest(
            new ByteArrayInputStream(
                GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate())));

    assertFalse(outputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertEquals(GridGrindProtocolCatalog.catalog(), catalog);
    assertEquals(GridGrindProtocolCatalog.requestTemplate(), template);
  }

  private static IllegalArgumentException invokeInvalidPayload(JacksonException exception) {
    return GridGrindJson.invalidPayload(exception);
  }

  /** Input stream wrapper that records whether caller-owned close was triggered. */
  private static final class TrackingInputStream extends InputStream {
    private final ByteArrayInputStream delegate;
    private boolean closed;

    private TrackingInputStream(byte[] bytes) {
      this.delegate = new ByteArrayInputStream(bytes);
    }

    @Override
    public int read() {
      return delegate.read();
    }

    @Override
    public int read(byte[] buffer, int offset, int length) {
      return delegate.read(buffer, offset, length);
    }

    @Override
    public void close() {
      closed = true;
    }
  }

  /** Synthetic Jackson exception used to cover wrapped validation-cause wording. */
  private static final class WrappedJacksonException extends JacksonException {
    private static final long serialVersionUID = 1L;

    private WrappedJacksonException(String message, Throwable cause) {
      super(message, cause);
    }

    @Override
    public TokenStreamLocation getLocation() {
      return null;
    }

    @Override
    public String getOriginalMessage() {
      return getMessage();
    }

    @Override
    public Object processor() {
      return null;
    }
  }
}
