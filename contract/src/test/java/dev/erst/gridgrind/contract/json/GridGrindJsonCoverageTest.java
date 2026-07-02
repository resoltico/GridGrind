package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.time.DateTimeException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ObjectNode;

/** Additional parser-wording and invalid-payload coverage for the shared JSON codec. */
class GridGrindJsonCoverageTest {
  @Test
  void readsResponsesAndCatalogsFromStreamsWithoutClosingThem() throws IOException {
    GridGrindResponse response =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.WorkbookInspectionResult
                    .WorkbookSummaryResult(
                    "summary", new WorkbookSummary.Empty(0, List.of(), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    RequestDoctorReport doctorReport =
        RequestDoctorReport.warnings(
            new RequestDoctorReport.Summary(
                "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 1, 1, 0, 0),
            List.of(new RequestWarning(0, "step-1", "SET_CELL", "warning")));

    try (TrackingInputStream responseStream =
            new TrackingInputStream(GridGrindJsonOutput.writeResponseBytes(response));
        TrackingInputStream catalogStream =
            new TrackingInputStream(GridGrindJsonOutput.writeProtocolCatalogBytes(catalog));
        TrackingInputStream doctorReportStream =
            new TrackingInputStream(
                GridGrindJsonOutput.writeRequestDoctorReportBytes(doctorReport))) {
      assertEquals(response, GridGrindJson.readResponse(responseStream));
      assertEquals(catalog, GridGrindJson.readProtocolCatalog(catalogStream));
      assertEquals(doctorReport, GridGrindJson.readRequestDoctorReport(doctorReportStream));
      assertEquals(
          doctorReport,
          GridGrindJson.readRequestDoctorReport(
              GridGrindJsonOutput.writeRequestDoctorReportBytes(doctorReport)));
      assertFalse(responseStream.closed);
      assertFalse(catalogStream.closed);
      assertFalse(doctorReportStream.closed);
    }
  }

  @Test
  void invalidResponseAndCatalogStreamsSurfaceInvalidJsonWithoutClosingCallerStreams()
      throws IOException {
    try (TrackingInputStream responseStream =
            new TrackingInputStream("{".getBytes(StandardCharsets.UTF_8));
        TrackingInputStream catalogStream =
            new TrackingInputStream("{".getBytes(StandardCharsets.UTF_8));
        TrackingInputStream doctorReportStream =
            new TrackingInputStream("{".getBytes(StandardCharsets.UTF_8))) {
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class, () -> GridGrindJson.readResponse(responseStream)));
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class, () -> GridGrindJson.readProtocolCatalog(catalogStream)));
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class,
              () -> GridGrindJson.readRequestDoctorReport(doctorReportStream)));
      assertFalse(responseStream.closed);
      assertFalse(catalogStream.closed);
      assertFalse(doctorReportStream.closed);
    }
  }

  @Test
  void invalidRequestBytesSurfaceInvalidJson() {
    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readRequest("{".getBytes(StandardCharsets.UTF_8))));
  }

  @Test
  void requestTreeRendersTheWireObjectWithoutIOLayering() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan explicitRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(ExecutionModeInput.eventRead()),
            new FormulaEnvironmentInput(
                List.of(new FormulaExternalWorkbookInput("rates.xlsx", "tmp/rates.xlsx")),
                FormulaMissingWorkbookPolicy.USE_CACHED_VALUE,
                List.of()),
            List.of());

    ObjectNode requestTree = GridGrindJsonOutput.requestTree(request);
    ObjectNode explicitRequestTree = GridGrindJsonOutput.requestTree(explicitRequest);

    assertEquals("V1", requestTree.path("protocolVersion").stringValue());
    assertEquals("NEW", requestTree.path("source").path("type").stringValue());
    assertTrue(requestTree.path("steps").isArray());
    assertFalse(requestTree.has("execution"));
    assertFalse(requestTree.has("formulaEnvironment"));
    assertEquals(
        "EVENT_READ",
        explicitRequestTree.path("execution").path("mode").path("type").stringValue());
    assertEquals(
        "USE_CACHED_VALUE",
        explicitRequestTree.path("formulaEnvironment").path("missingWorkbookPolicy").stringValue());
    assertEquals(
        "request must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindJsonOutput.requestTree(null))
            .getMessage());
  }

  @Test
  void readRequestFromJsonStringDecodesWithoutATransportLayer() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    String requestJson = GridGrindJsonOutput.requestTree(request).toString();

    assertEquals(request, GridGrindJson.readRequest(requestJson));
    assertEquals(
        "json must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindJson.readRequest((String) null))
            .getMessage());
    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(InvalidJsonException.class, () -> GridGrindJson.readRequest("{")));
    InvalidRequestShapeException explicitNull =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": null,
                      "steps": [ ]
                    }
                    """));
    assertEquals(
        "Field 'execution' must be omitted when absent; explicit null is not accepted.",
        explicitNull.getMessage());
  }

  @Test
  void invalidRequestDoctorReportBytesSurfaceInvalidJson() {
    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readRequestDoctorReport("{".getBytes(StandardCharsets.UTF_8))));
  }

  @Test
  void emptyJsonDocumentsSurfaceInvalidJsonAcrossByteAndStreamReads() {
    assertEquals(
        "Invalid JSON payload",
        assertThrows(InvalidJsonException.class, () -> GridGrindJson.readRequest(new byte[0]))
            .getMessage());
    assertEquals(
        "Invalid JSON payload",
        assertThrows(
                InvalidJsonException.class,
                () -> GridGrindJson.readResponse(new ByteArrayInputStream(new byte[0])))
            .getMessage());
  }

  @Test
  void rejectsTopLevelAndArrayNullRequestPayloads() {
    assertEquals(
        "JSON payload must not be null",
        assertThrows(
                InvalidRequestShapeException.class,
                () -> GridGrindJson.readRequest("null".getBytes(StandardCharsets.UTF_8)))
            .getMessage());
    assertEquals(
        "JSON payload must not be null",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        new ByteArrayInputStream("null".getBytes(StandardCharsets.UTF_8))))
            .getMessage());
    assertEquals(
        "Field 'steps[0]' must be omitted when absent; explicit null is not accepted.",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        """
                        {
                          "protocolVersion": "V1",
                          "source": { "type": "NEW" },
                          "persistence": { "type": "NONE" },
                          "execution": {
                            "mode": {"type": "FULL_XSSF"},
                            "journal": { "level": "NORMAL" },
                            "calculation": {
                              "strategy": { "type": "DO_NOT_CALCULATE" },
                              "markRecalculateOnOpen": false
                            }
                          },
                          "formulaEnvironment": {
                            "externalWorkbooks": [],
                            "missingWorkbookPolicy": "ERROR",
                            "udfToolpacks": []
                          },
                          "steps": [null]
                        }
                        """
                            .getBytes(StandardCharsets.UTF_8)))
            .getMessage());
  }

  @Test
  void rejectsExplicitNullPlaceholdersAcrossNonRequestWireReads() throws IOException {
    GridGrindResponse response =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.WorkbookInspectionResult
                    .WorkbookSummaryResult(
                    "summary", new WorkbookSummary.Empty(0, List.of(), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    RequestDoctorReport doctorReport =
        RequestDoctorReport.clean(
            new RequestDoctorReport.Summary(
                "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0));

    assertEquals(
        "Field 'warnings' must be omitted when absent; explicit null is not accepted.",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readResponse(
                        new ByteArrayInputStream(
                            withTopLevelNull(
                                GridGrindJsonOutput.writeResponseBytes(response), "warnings"))))
            .getMessage());
    assertEquals(
        "Field 'warnings' must be omitted when absent; explicit null is not accepted.",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readResponse(
                        withTopLevelNull(
                            GridGrindJsonOutput.writeResponseBytes(response), "warnings")))
            .getMessage());
    assertEquals(
        "Field 'plainTypes' must be omitted when absent; explicit null is not accepted.",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readProtocolCatalog(
                        withTopLevelNull(
                            GridGrindJsonOutput.writeProtocolCatalogBytes(catalog), "plainTypes")))
            .getMessage());
    assertEquals(
        "Field 'warnings' must be omitted when absent; explicit null is not accepted.",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequestDoctorReport(
                        withTopLevelNull(
                            GridGrindJsonOutput.writeRequestDoctorReportBytes(doctorReport),
                            "warnings")))
            .getMessage());
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
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.readRequestDoctorReport((InputStream) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindJson.readProtocolCatalog((byte[]) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.readRequestDoctorReport((byte[]) null))
            .getMessage());
    assertEquals(
        "request must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindJsonOutput.writeRequestBytes(null))
            .getMessage());
    assertEquals(
        "response must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindJsonOutput.writeResponseBytes(null))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJsonOutput.writeResponse(
                        null,
                        GridGrindResponses.failure(
                            new GridGrindProblemDetail.Problem(
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
                                new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces
                                        .CliArgument.named("--request")),
                                java.util.Optional.empty(),
                                List.of())),
                        false))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJsonOutput.writeProtocolCatalog(
                        null, GridGrindProtocolCatalog.catalog(), false))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJsonOutput.writeRequestDoctorReport(
                        null,
                        RequestDoctorReport.clean(
                            new RequestDoctorReport.Summary(
                                "NEW",
                                "NONE",
                                "FULL_XSSF",
                                "DO_NOT_CALCULATE",
                                false,
                                false,
                                0,
                                0,
                                0,
                                0)),
                        false))
            .getMessage());
  }

  @Test
  void helperMethodsCoverFallbackAndLocationEdgeCases() throws IOException {
    JsonFactory jsonFactory = new JsonFactory();
    MismatchedInputException floatingPointWithoutPath =
        MismatchedInputException.from(
            utf8Parser(jsonFactory, "2.5"),
            Integer.class,
            "Cannot coerce Floating-point value (2.5) to `int` value"
                + " (but could if coercion was enabled using `CoercionConfig`)");
    MismatchedInputException floatingPointWithIndex =
        (MismatchedInputException)
            MismatchedInputException.from(
                    utf8Parser(jsonFactory, "2.5"),
                    Integer.class,
                    "Cannot coerce Floating-point value (2.5) to `int` value"
                        + " (but could if coercion was enabled using `CoercionConfig`)")
                .prependPath(new Object(), 0);
    MismatchedInputException floatingPointWithNestedPath =
        (MismatchedInputException)
            MismatchedInputException.from(
                    utf8Parser(jsonFactory, "2.5"),
                    Integer.class,
                    "Cannot coerce Floating-point value (2.5) to `int` value"
                        + " (but could if coercion was enabled using `CoercionConfig`)")
                .prependPath(new Object(), 1)
                .prependPath(new Object(), "bar")
                .prependPath(new Object(), 0)
                .prependPath(new Object(), "items");

    assertEquals(
        "JSON value must be an integer value",
        GridGrindJson.mismatchedInputMessage(floatingPointWithoutPath));
    assertEquals(
        "JSON value at '[0]' must be an integer value",
        GridGrindJson.mismatchedInputMessage(floatingPointWithIndex));
    assertEquals(
        "JSON value at 'items[0].bar[1]' must be an integer value",
        GridGrindJson.mismatchedInputMessage(floatingPointWithNestedPath));
    assertEquals(
        "fieldName must not be null",
        GridGrindJson.message(new NullPointerException("fieldName must not be null")));
    assertEquals(
        "steps[0].target must not be null",
        GridGrindJson.message(new NullPointerException("steps[0].target must not be null")));
    assertEquals(
        "Cannot deserialize value of type `x` from String",
        GridGrindJson.message(
            new IllegalArgumentException("Cannot deserialize value of type `x` from String")));
    assertEquals("bad", invokeInvalidPayload(new StreamConstraintsException("bad")).getMessage());
    assertEquals(
        "wrapper",
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
        "Missing required creator property 'fieldName'",
        GridGrindJson.message(
            new IllegalStateException("Missing required creator property 'fieldName'")));
    assertEquals(
        "missing type id property 'type'",
        GridGrindJson.message(new IllegalStateException("missing type id property 'type'")));
    assertEquals(
        "Invalid JSON payload",
        GridGrindJson.mismatchedInputMessage(
            MismatchedInputException.from(
                utf8Parser(jsonFactory, "null"), Integer.class, (String) null)));
    assertEquals(
        Optional.empty(), GridGrindJson.jsonLine(new TokenStreamLocation(null, 0L, 0L, 0, 9)));
    assertEquals(
        Optional.empty(), GridGrindJson.jsonColumn(new TokenStreamLocation(null, 0L, 0L, 4, 0)));
    assertEquals(
        "Cannot coerce value to `boolean`"
            + " (but could if coercion was enabled using `CoercionConfig`)",
        GridGrindJson.cleanJacksonMessage(
            "Cannot coerce value to `boolean`"
                + " (but could if coercion was enabled using `CoercionConfig`)"));
  }

  @Test
  void typeEntryAndRequestReadersStayDeterministicThroughPublicRoundTrip() throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    ByteArrayOutputStream doctorReportOutputStream = new ByteArrayOutputStream();
    TypeEntry entry = GridGrindProtocolCatalog.entryFor("GET_CELLS").orElseThrow();
    RequestDoctorReport doctorReport =
        RequestDoctorReport.clean(
            new RequestDoctorReport.Summary(
                "NEW", "SAVE_AS", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0));
    GridGrindJsonOutput.writeTypeEntry(outputStream, entry);
    GridGrindJsonOutput.writeRequestDoctorReport(doctorReportOutputStream, doctorReport, false);
    Catalog catalog =
        GridGrindJson.readProtocolCatalog(
            new ByteArrayInputStream(
                GridGrindJsonOutput.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog())));
    RequestDoctorReport decodedDoctorReport =
        GridGrindJson.readRequestDoctorReport(
            new ByteArrayInputStream(
                GridGrindJsonOutput.writeRequestDoctorReportBytes(doctorReport)));
    WorkbookPlan template =
        GridGrindJson.readRequest(
            new ByteArrayInputStream(
                GridGrindJsonOutput.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate())));

    assertFalse(outputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertFalse(doctorReportOutputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertEquals(GridGrindProtocolCatalog.catalog(), catalog);
    assertEquals(doctorReport, decodedDoctorReport);
    assertEquals(GridGrindProtocolCatalog.requestTemplate(), template);
  }

  @Test
  void streamWriteApisMatchByteArraySerializationsForLargePayloads() throws IOException {
    String largeText = "x".repeat(200_000);
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
                        {
                          "protocolVersion": "V1",
                          "source": { "type": "NEW" },
                          "persistence": { "type": "NONE" },
                          "execution": {
                            "mode": {"type": "FULL_XSSF"},
                            "journal": { "level": "NORMAL" },
                            "calculation": {
                              "strategy": { "type": "DO_NOT_CALCULATE" },
                              "markRecalculateOnOpen": false
                            }
                          },
                          "formulaEnvironment": {
                            "externalWorkbooks": [],
                            "missingWorkbookPolicy": "ERROR",
                            "udfToolpacks": []
                          },
                          "steps": [
                            {
                          "stepId": "set-owner",
                          "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                          "action": {
                            "type": "SET_CELL",
                            "value": {
                              "type": "TEXT",
                              "source": { "type": "INLINE", "text": "%s" }
                            }
                          }
                        }
                      ]
                    }
                    """
                .formatted(largeText)
                .getBytes(StandardCharsets.UTF_8));
    GridGrindResponse response =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.WorkbookInspectionResult
                    .WorkbookSummaryResult(
                    "summary", new WorkbookSummary.Empty(0, List.of(), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    assertStreamSerializationMatchesBytes(
        GridGrindJsonOutput.writeRequestBytes(request),
        out -> GridGrindJsonOutput.writeRequest(out, request, false));
    assertStreamSerializationMatchesBytes(
        GridGrindJsonOutput.writeResponseBytes(response),
        out -> GridGrindJsonOutput.writeResponse(out, response, false));
    assertStreamSerializationMatchesBytes(
        GridGrindJsonOutput.writeProtocolCatalogBytes(catalog),
        out -> GridGrindJsonOutput.writeProtocolCatalog(out, catalog, false));
  }

  @Test
  void catalogLookupResultPrependsProtocolVersionToValueFields() throws IOException {
    TypeEntry entry = GridGrindProtocolCatalog.entryFor("GET_CELLS").orElseThrow();
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    GridGrindJsonOutput.writeCatalogLookupResult(outputStream, GridGrindProtocolVersion.V1, entry);
    String json = outputStream.toString(StandardCharsets.UTF_8);
    assertTrue(json.indexOf("\"protocolVersion\"") < json.indexOf("\"id\""));
    assertTrue(json.contains("\"GET_CELLS\""));
    assertEquals(
        "protocolVersion must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJsonOutput.writeCatalogLookupResult(
                        new ByteArrayOutputStream(), null, entry))
            .getMessage());
    assertEquals(
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJsonOutput.writeCatalogLookupResult(
                        new ByteArrayOutputStream(), GridGrindProtocolVersion.V1, null))
            .getMessage());
  }

  @Test
  void discoverySerializersOmitExplicitNullProperties() throws IOException {
    String catalogJson =
        new String(
            GridGrindJsonOutput.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()),
            StandardCharsets.UTF_8);
    ByteArrayOutputStream typeEntryOutput = new ByteArrayOutputStream();
    GridGrindJsonOutput.writeTypeEntry(
        typeEntryOutput, GridGrindProtocolCatalog.entryFor("EXPECT_TABLE_PRESENT").orElseThrow());
    String typeEntryJson = typeEntryOutput.toString(StandardCharsets.UTF_8);

    assertFalse(catalogJson.contains(": null"));
    assertFalse(typeEntryJson.contains(": null"));
    assertFalse(typeEntryJson.contains("targetSelectorRule"));
  }

  @Test
  void requestAndResponseSerializersOmitExplicitNullProperties() throws IOException {
    WorkbookPlan request = GridGrindProtocolCatalog.requestTemplate();
    GridGrindResponse response =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.WorkbookInspectionResult
                    .WorkbookSummaryResult(
                    "summary", new WorkbookSummary.Empty(0, List.of(), 0, false))));

    String requestJson =
        new String(GridGrindJsonOutput.writeRequestBytes(request), StandardCharsets.UTF_8);
    String responseJson =
        new String(GridGrindJsonOutput.writeResponseBytes(response), StandardCharsets.UTF_8);

    assertFalse(requestJson.contains(": null"));
    assertFalse(responseJson.contains(": null"));
    assertFalse(requestJson.contains("\"execution\""));
    assertFalse(requestJson.contains("\"formulaEnvironment\""));
  }

  @Test
  void compactReadbackResponsesStayWithinTheM3PayloadBudgets() throws IOException {
    GridGrindResponse cellsResponse =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.SheetInspectionResult.CellsResult(
                    "cells",
                    "Budget",
                    List.of(
                        new CellReport.TextReport(
                            "A1",
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of("Ada"),
                            Optional.empty())))));
    GridGrindResponse sparseWindowResponse =
        GridGrindResponses.success(
            List.of(),
            List.of(),
            List.of(
                new dev.erst.gridgrind.contract.query.SheetInspectionResult.WindowResult(
                    "window",
                    new WindowReport.Sparse(
                        "Budget",
                        "A1",
                        new WindowDimensionsReport(50, 50),
                        List.of(
                            new CellReport.TextReport(
                                "A1",
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.of("Ada"),
                                Optional.empty()))))));

    byte[] cellsResponseBytes = GridGrindJsonOutput.writeResponseBytes(cellsResponse);
    byte[] sparseWindowResponseBytes = GridGrindJsonOutput.writeResponseBytes(sparseWindowResponse);

    assertTrue(
        cellsResponseBytes.length < 1316,
        () ->
            "default one-cell readback should stay compact but serialized to "
                + cellsResponseBytes.length
                + " bytes");
    assertTrue(
        sparseWindowResponseBytes.length < 4096,
        () ->
            "sparse 50x50 near-empty window should stay in kilobytes but serialized to "
                + sparseWindowResponseBytes.length
                + " bytes");
  }

  @Test
  void rejectsExplicitNullInDeepNestedRequestFieldWithFullDottedPath() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": {"type": "FULL_XSSF"},
                        "journal": { "level": "NORMAL" },
                        "calculation": {
                          "strategy": { "type": "DO_NOT_CALCULATE" },
                          "markRecalculateOnOpen": null
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    assertEquals(
        "Field 'execution.calculation.markRecalculateOnOpen' must be omitted when absent; explicit"
            + " null is not accepted.",
        exception.getMessage());
    assertEquals(
        Optional.of("execution.calculation.markRecalculateOnOpen"),
        exception.jsonLocation().jsonPath());
  }

  @Test
  void classifiesOnlyExplicitNullChecksAsValidationCauses() {
    NullPointerException explicitNull = new NullPointerException("field must not be null");
    NullPointerException jvmNull = new NullPointerException("Cannot invoke method on null");
    NullPointerException nullMessageNull = new NullPointerException();

    assertInstanceOf(
        InvalidJsonException.class,
        invokeInvalidPayload(new WrappedJacksonException("wrapper", explicitNull)),
        "Synthetic helper calls without structural intake context stay in the invalid-JSON lane.");
    assertInstanceOf(
        InvalidJsonException.class,
        invokeInvalidPayload(new WrappedJacksonException("wrapper", jvmNull)),
        "JVM NPE without 'must not be null' message should not be treated as a validation error");
    assertInstanceOf(
        InvalidJsonException.class,
        invokeInvalidPayload(new WrappedJacksonException("wrapper", nullMessageNull)),
        "NPE with null message should not be treated as a validation error");
  }

  @Test
  void prefersMissingRequiredCreatorMessagesOverExplicitNullValidationCauses() {
    InvalidJsonException failure =
        assertInstanceOf(
            InvalidJsonException.class,
            invokeInvalidPayload(
                new WrappedJacksonException(
                    "Missing required creator property 'protocolVersion'",
                    new NullPointerException("protocolVersion must not be null"))));

    assertEquals("Missing required creator property 'protocolVersion'", failure.getMessage());
  }

  private static IllegalArgumentException invokeInvalidPayload(JacksonException exception) {
    return GridGrindJson.invalidPayload(exception);
  }

  private static tools.jackson.core.JsonParser utf8Parser(JsonFactory jsonFactory, String json)
      throws IOException {
    return jsonFactory.createParser(
        tools.jackson.core.ObjectReadContext.empty(),
        new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)));
  }

  private static void assertStreamSerializationMatchesBytes(byte[] expected, StreamWriter writer)
      throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    writer.write(outputStream);
    assertArrayEquals(expected, outputStream.toByteArray());
  }

  private static byte[] withTopLevelNull(byte[] bytes, String fieldName) throws IOException {
    JsonMapper mapper = JsonMapper.builder().build();
    ObjectNode node = (ObjectNode) mapper.readTree(bytes);
    node.putNull(fieldName);
    return mapper.writeValueAsBytes(node);
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

  /** One streamed JSON writer shape used to prove output-stream APIs avoid buffering drift. */
  @FunctionalInterface
  @SuppressWarnings("NotJavadoc")
  private interface StreamWriter {
    void write(OutputStream outputStream) throws IOException;
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
