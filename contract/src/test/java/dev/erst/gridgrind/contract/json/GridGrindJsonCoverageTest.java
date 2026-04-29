package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GoalPlanReport;
import dev.erst.gridgrind.contract.catalog.GridGrindGoalPlanner;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskPlanner;
import dev.erst.gridgrind.contract.catalog.TaskCatalog;
import dev.erst.gridgrind.contract.catalog.TaskPlanTemplate;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
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
                new dev.erst.gridgrind.contract.query.InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    TaskCatalog taskCatalog = GridGrindTaskCatalog.catalog();
    TaskPlanTemplate taskPlanTemplate = GridGrindTaskPlanner.templateFor("DASHBOARD");
    GoalPlanReport goalPlanReport =
        GridGrindGoalPlanner.reportFor("Create a monthly sales dashboard with charts");
    RequestDoctorReport doctorReport =
        RequestDoctorReport.warnings(
            new RequestDoctorReport.Summary(
                "NEW",
                "NONE",
                "FULL_XSSF",
                "FULL_XSSF",
                "DO_NOT_CALCULATE",
                false,
                false,
                1,
                1,
                0,
                0),
            List.of(new RequestWarning(0, "step-1", "SET_CELL", "warning")));

    try (TrackingInputStream responseStream =
            new TrackingInputStream(GridGrindJson.writeResponseBytes(response));
        TrackingInputStream catalogStream =
            new TrackingInputStream(GridGrindJson.writeProtocolCatalogBytes(catalog));
        TrackingInputStream taskCatalogStream =
            new TrackingInputStream(GridGrindJson.writeTaskCatalogBytes(taskCatalog));
        TrackingInputStream taskPlanStream =
            new TrackingInputStream(GridGrindJson.writeTaskPlanTemplateBytes(taskPlanTemplate));
        TrackingInputStream goalPlanStream =
            new TrackingInputStream(GridGrindJson.writeGoalPlanReportBytes(goalPlanReport));
        TrackingInputStream doctorReportStream =
            new TrackingInputStream(GridGrindJson.writeRequestDoctorReportBytes(doctorReport))) {
      assertEquals(response, GridGrindJson.readResponse(responseStream));
      assertEquals(catalog, GridGrindJson.readProtocolCatalog(catalogStream));
      assertEquals(taskCatalog, GridGrindJson.readTaskCatalog(taskCatalogStream));
      assertEquals(taskPlanTemplate, GridGrindJson.readTaskPlanTemplate(taskPlanStream));
      assertEquals(goalPlanReport, GridGrindJson.readGoalPlanReport(goalPlanStream));
      assertEquals(doctorReport, GridGrindJson.readRequestDoctorReport(doctorReportStream));
      assertFalse(responseStream.closed);
      assertFalse(catalogStream.closed);
      assertFalse(taskCatalogStream.closed);
      assertFalse(taskPlanStream.closed);
      assertFalse(goalPlanStream.closed);
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
        TrackingInputStream taskCatalogStream =
            new TrackingInputStream("{".getBytes(StandardCharsets.UTF_8));
        TrackingInputStream taskPlanStream =
            new TrackingInputStream("{".getBytes(StandardCharsets.UTF_8));
        TrackingInputStream goalPlanStream =
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
              InvalidJsonException.class, () -> GridGrindJson.readTaskCatalog(taskCatalogStream)));
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class,
              () -> GridGrindJson.readTaskPlanTemplate(taskPlanStream)));
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class, () -> GridGrindJson.readGoalPlanReport(goalPlanStream)));
      assertInstanceOf(
          InvalidJsonException.class,
          assertThrows(
              InvalidJsonException.class,
              () -> GridGrindJson.readRequestDoctorReport(doctorReportStream)));
      assertFalse(responseStream.closed);
      assertFalse(catalogStream.closed);
      assertFalse(taskCatalogStream.closed);
      assertFalse(taskPlanStream.closed);
      assertFalse(goalPlanStream.closed);
      assertFalse(doctorReportStream.closed);
    }
  }

  @Test
  void invalidTaskCatalogBytesSurfaceInvalidJson() {
    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readTaskCatalog("{".getBytes(StandardCharsets.UTF_8))));
  }

  @Test
  void invalidTaskPlanTemplateBytesSurfaceInvalidJson() {
    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readTaskPlanTemplate("{".getBytes(StandardCharsets.UTF_8))));
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
  void invalidGoalPlanReportBytesSurfaceInvalidJson() {
    assertInstanceOf(
        InvalidJsonException.class,
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readGoalPlanReport("{".getBytes(StandardCharsets.UTF_8))));
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
        "problem: <root> must not be null",
        assertThrows(
                InvalidRequestException.class,
                () -> GridGrindJson.readRequest("null".getBytes(StandardCharsets.UTF_8)))
            .getMessage());
    assertEquals(
        "problem: <root> must not be null",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readRequest(
                        new ByteArrayInputStream("null".getBytes(StandardCharsets.UTF_8))))
            .getMessage());
    assertEquals(
        "Missing required field 'steps[0]'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readRequest(
                        """
                        {
                          "source": { "type": "NEW" },
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
                new dev.erst.gridgrind.contract.query.InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    TaskCatalog taskCatalog = GridGrindTaskCatalog.catalog();
    TaskPlanTemplate taskPlanTemplate = GridGrindTaskPlanner.templateFor("DASHBOARD");
    GoalPlanReport goalPlanReport =
        GridGrindGoalPlanner.reportFor("Create a monthly sales dashboard with charts");
    RequestDoctorReport doctorReport =
        RequestDoctorReport.clean(
            new RequestDoctorReport.Summary(
                "NEW",
                "NONE",
                "FULL_XSSF",
                "FULL_XSSF",
                "DO_NOT_CALCULATE",
                false,
                false,
                0,
                0,
                0,
                0));

    assertEquals(
        "Missing required field 'warnings'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readResponse(
                        new ByteArrayInputStream(
                            withTopLevelNull(
                                GridGrindJson.writeResponseBytes(response), "warnings"))))
            .getMessage());
    assertEquals(
        "Missing required field 'warnings'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readResponse(
                        withTopLevelNull(GridGrindJson.writeResponseBytes(response), "warnings")))
            .getMessage());
    assertEquals(
        "Missing required field 'shippedExamples'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readProtocolCatalog(
                        withTopLevelNull(
                            GridGrindJson.writeProtocolCatalogBytes(catalog), "shippedExamples")))
            .getMessage());
    assertEquals(
        "Missing required field 'tasks'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readTaskCatalog(
                        withTopLevelNull(
                            GridGrindJson.writeTaskCatalogBytes(taskCatalog), "tasks")))
            .getMessage());
    assertEquals(
        "Missing required field 'authoringNotes'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readTaskPlanTemplate(
                        withTopLevelNull(
                            GridGrindJson.writeTaskPlanTemplateBytes(taskPlanTemplate),
                            "authoringNotes")))
            .getMessage());
    assertEquals(
        "Missing required field 'candidates'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readGoalPlanReport(
                        withTopLevelNull(
                            GridGrindJson.writeGoalPlanReportBytes(goalPlanReport), "candidates")))
            .getMessage());
    assertEquals(
        "Missing required field 'warnings'",
        assertThrows(
                InvalidRequestException.class,
                () ->
                    GridGrindJson.readRequestDoctorReport(
                        withTopLevelNull(
                            GridGrindJson.writeRequestDoctorReportBytes(doctorReport), "warnings")))
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
                NullPointerException.class, () -> GridGrindJson.readTaskCatalog((InputStream) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.readTaskPlanTemplate((InputStream) null))
            .getMessage());
    assertEquals(
        "inputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.readGoalPlanReport((InputStream) null))
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
        assertThrows(NullPointerException.class, () -> GridGrindJson.readTaskCatalog((byte[]) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindJson.readTaskPlanTemplate((byte[]) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class, () -> GridGrindJson.readGoalPlanReport((byte[]) null))
            .getMessage());
    assertEquals(
        "bytes must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.readRequestDoctorReport((byte[]) null))
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
                        GridGrindResponses.failure(
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
                                new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                                    dev.erst.gridgrind.contract.dto.ProblemContext.CliArgument
                                        .named("--request")),
                                java.util.Optional.empty(),
                                List.of()))))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.writeProtocolCatalog(null, GridGrindProtocolCatalog.catalog()))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.writeTaskCatalog(null, GridGrindTaskCatalog.catalog()))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJson.writeTaskPlanTemplate(
                        null, GridGrindTaskPlanner.templateFor("DASHBOARD")))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJson.writeGoalPlanReport(
                        null,
                        GridGrindGoalPlanner.reportFor(
                            "Create a monthly sales dashboard with charts")))
            .getMessage());
    assertEquals(
        "outputStream must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    GridGrindJson.writeRequestDoctorReport(
                        null,
                        RequestDoctorReport.clean(
                            new RequestDoctorReport.Summary(
                                "NEW",
                                "NONE",
                                "FULL_XSSF",
                                "FULL_XSSF",
                                "DO_NOT_CALCULATE",
                                false,
                                false,
                                0,
                                0,
                                0,
                                0))))
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
        "Missing required field 'steps[0].target'",
        GridGrindJson.message(
            new IllegalArgumentException("problem: steps[0].target must not be null")));
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
    assertEquals(
        Optional.empty(), GridGrindJson.jsonLine(new TokenStreamLocation(null, 0L, 0L, 0, 9)));
    assertEquals(
        Optional.empty(), GridGrindJson.jsonColumn(new TokenStreamLocation(null, 0L, 0L, 4, 0)));
  }

  @Test
  void typeEntryAndRequestReadersStayDeterministicThroughPublicRoundTrip() throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    ByteArrayOutputStream taskOutputStream = new ByteArrayOutputStream();
    ByteArrayOutputStream taskCatalogOutputStream = new ByteArrayOutputStream();
    ByteArrayOutputStream taskPlanOutputStream = new ByteArrayOutputStream();
    ByteArrayOutputStream goalPlanOutputStream = new ByteArrayOutputStream();
    ByteArrayOutputStream doctorReportOutputStream = new ByteArrayOutputStream();
    TypeEntry entry = GridGrindProtocolCatalog.entryFor("GET_CELLS").orElseThrow();
    var taskEntry = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskCatalog taskCatalog = GridGrindTaskCatalog.catalog();
    TaskPlanTemplate taskPlanTemplate = GridGrindTaskPlanner.templateFor("DASHBOARD");
    GoalPlanReport goalPlanReport =
        GridGrindGoalPlanner.reportFor("Create a monthly sales dashboard with charts");
    RequestDoctorReport doctorReport =
        RequestDoctorReport.clean(
            new RequestDoctorReport.Summary(
                "NEW",
                "SAVE_AS",
                "FULL_XSSF",
                "FULL_XSSF",
                "DO_NOT_CALCULATE",
                false,
                false,
                0,
                0,
                0,
                0));
    GridGrindJson.writeTypeEntry(outputStream, entry);
    GridGrindJson.writeTaskEntry(taskOutputStream, taskEntry);
    GridGrindJson.writeTaskCatalog(taskCatalogOutputStream, taskCatalog);
    GridGrindJson.writeTaskPlanTemplate(taskPlanOutputStream, taskPlanTemplate);
    GridGrindJson.writeGoalPlanReport(goalPlanOutputStream, goalPlanReport);
    GridGrindJson.writeRequestDoctorReport(doctorReportOutputStream, doctorReport);
    Catalog catalog =
        GridGrindJson.readProtocolCatalog(
            new ByteArrayInputStream(
                GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog())));
    TaskCatalog decodedTaskCatalog =
        GridGrindJson.readTaskCatalog(
            new ByteArrayInputStream(GridGrindJson.writeTaskCatalogBytes(taskCatalog)));
    TaskPlanTemplate decodedTaskPlanTemplate =
        GridGrindJson.readTaskPlanTemplate(
            new ByteArrayInputStream(GridGrindJson.writeTaskPlanTemplateBytes(taskPlanTemplate)));
    GoalPlanReport decodedGoalPlanReport =
        GridGrindJson.readGoalPlanReport(
            new ByteArrayInputStream(GridGrindJson.writeGoalPlanReportBytes(goalPlanReport)));
    RequestDoctorReport decodedDoctorReport =
        GridGrindJson.readRequestDoctorReport(
            new ByteArrayInputStream(GridGrindJson.writeRequestDoctorReportBytes(doctorReport)));
    WorkbookPlan template =
        GridGrindJson.readRequest(
            new ByteArrayInputStream(
                GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate())));

    assertFalse(outputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertFalse(taskOutputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertFalse(taskCatalogOutputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertFalse(taskPlanOutputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertFalse(goalPlanOutputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertFalse(doctorReportOutputStream.toString(StandardCharsets.UTF_8).isBlank());
    assertEquals(GridGrindProtocolCatalog.catalog(), catalog);
    assertEquals(taskCatalog, decodedTaskCatalog);
    assertEquals(taskPlanTemplate, decodedTaskPlanTemplate);
    assertEquals(goalPlanReport, decodedGoalPlanReport);
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
                      "source": { "type": "NEW" },
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
                new dev.erst.gridgrind.contract.query.InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    assertStreamSerializationMatchesBytes(
        GridGrindJson.writeRequestBytes(request), out -> GridGrindJson.writeRequest(out, request));
    assertStreamSerializationMatchesBytes(
        GridGrindJson.writeResponseBytes(response),
        out -> GridGrindJson.writeResponse(out, response));
    assertStreamSerializationMatchesBytes(
        GridGrindJson.writeProtocolCatalogBytes(catalog),
        out -> GridGrindJson.writeProtocolCatalog(out, catalog));
  }

  @Test
  void discoverySerializersOmitExplicitNullProperties() throws IOException {
    String catalogJson =
        new String(
            GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog()),
            StandardCharsets.UTF_8);
    ByteArrayOutputStream typeEntryOutput = new ByteArrayOutputStream();
    GridGrindJson.writeTypeEntry(
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
                new dev.erst.gridgrind.contract.query.InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.Empty(0, List.of(), 0, false))));

    String requestJson =
        new String(GridGrindJson.writeRequestBytes(request), StandardCharsets.UTF_8);
    String responseJson =
        new String(GridGrindJson.writeResponseBytes(response), StandardCharsets.UTF_8);

    assertFalse(requestJson.contains(": null"));
    assertFalse(responseJson.contains(": null"));
  }

  private static IllegalArgumentException invokeInvalidPayload(JacksonException exception) {
    return GridGrindJson.invalidPayload(exception);
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
