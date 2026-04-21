package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.databind.exc.InvalidTypeIdException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;

/** Tests for JSON serialization, parser wording, and the step-based wire shape. */
@SuppressWarnings("NotJavadoc")
class GridGrindJsonTest {
  @Test
  void defaultsRequestProtocolVersionDuringJsonRead() throws IOException {
    WorkbookPlan plan =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(GridGrindProtocolVersion.V1, plan.protocolVersion());
  }

  @Test
  @SuppressWarnings("StringConcatToTextBlock")
  void roundTripsRequestsResponsesAndCatalogs() throws IOException {
    WorkbookPlan request =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.FULL_XSSF, ExecutionModeInput.WriteMode.FULL_XSSF),
            null,
            List.of(
                new MutationStep(
                    "set-owner",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(new CellInput.Text(text("Owner")))),
                new AssertionStep(
                    "assert-owner",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));
    GridGrindResponse response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            List.of(new RequestWarning(0, "set-owner", "SET_CELL", "warning")),
            List.of(new AssertionResult("assert-owner", "EXPECT_CELL_VALUE")),
            List.of(
                new InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();

    assertEquals(request, GridGrindJson.readRequest(GridGrindJson.writeRequestBytes(request)));
    assertEquals(response, GridGrindJson.readResponse(GridGrindJson.writeResponseBytes(response)));
    assertEquals(
        catalog,
        GridGrindJson.readProtocolCatalog(GridGrindJson.writeProtocolCatalogBytes(catalog)));
  }

  @Test
  void roundTripsResolveInputsAndCalculationFailureContexts() throws IOException {
    GridGrindResponse resolveInputsFailure =
        new GridGrindResponse.Failure(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.Problem(
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND
                    .category(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND
                    .recovery(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND.title(),
                "missing payload",
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND
                    .resolution(),
                new GridGrindResponse.ProblemContext.ResolveInputs(
                    "NEW", "NONE", "cell text", "missing.txt"),
                null,
                List.of()));
    GridGrindResponse calculationFailure =
        new GridGrindResponse.Failure(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.Problem(
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA,
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.category(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.recovery(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.title(),
                "bad formula",
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.resolution(),
                new GridGrindResponse.ProblemContext.ExecuteCalculation.Preflight(
                    "EXISTING", "SAVE_AS", "Ops", "B1", "SUM("),
                null,
                List.of()));

    assertEquals(
        resolveInputsFailure,
        GridGrindJson.readResponse(GridGrindJson.writeResponseBytes(resolveInputsFailure)));
    assertEquals(
        calculationFailure,
        GridGrindJson.readResponse(GridGrindJson.writeResponseBytes(calculationFailure)));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  @Test
  void readsCalculationPolicyWithoutExplicitOpenTimeRecalcFlagAndDoesNotLeakHelperFields()
      throws IOException {
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "execution": {
                "calculation": {
                  "strategy": { "type": "EVALUATE_ALL" }
                }
              },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    byte[] serialized = GridGrindJson.writeRequestBytes(request);
    String serializedJson = new String(serialized, StandardCharsets.UTF_8);

    assertFalse(serializedJson.contains("\"default\""));
    assertEquals(
        new ExecutionModeInput(
            ExecutionModeInput.ReadMode.FULL_XSSF, ExecutionModeInput.WriteMode.FULL_XSSF),
        request.effectiveExecutionMode());
    assertFalse(request.calculationPolicy().markRecalculateOnOpen());
    assertEquals(request, GridGrindJson.readRequest(serialized));
  }

  @Test
  void readsCanonicalStepEnvelopeWithoutOuterStepType() throws IOException {
    WorkbookPlan plan =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "steps": [
                {
                  "stepId": "set-owner",
                  "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                  "action": {
                    "type": "SET_CELL",
                    "value": {
                      "type": "TEXT",
                      "source": { "type": "INLINE", "text": "Owner" }
                    }
                  }
                },
                {
                  "stepId": "assert-owner",
                  "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                  "assertion": {
                    "type": "EXPECT_CELL_VALUE",
                    "expectedValue": { "type": "TEXT", "text": "Owner" }
                  }
                },
                {
                  "stepId": "summary",
                  "target": { "type": "CURRENT" },
                  "query": { "type": "GET_WORKBOOK_SUMMARY" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(3, plan.steps().size());
    assertInstanceOf(MutationStep.class, plan.steps().get(0));
    assertInstanceOf(AssertionStep.class, plan.steps().get(1));
    assertInstanceOf(InspectionStep.class, plan.steps().get(2));
  }

  @Test
  void readsRequestsFromInputStreamsWithoutClosingThem() throws IOException {
    try (TrackingInputStream inputStream =
        new TrackingInputStream(
            """
            {
              "source": { "type": "NEW" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8))) {
      WorkbookPlan request = GridGrindJson.readRequest(inputStream);

      assertEquals(GridGrindProtocolVersion.V1, request.protocolVersion());
      assertFalse(inputStream.closed);
    }
  }

  @Test
  void rejectsUnknownTypeValuesAndFractionalIntegersWithProductOwnedMessages() {
    InvalidRequestShapeException unknownType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "bad",
                          "target": { "type": "CURRENT" },
                          "query": { "type": "NO_SUCH_QUERY" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException fractionalInteger =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "window",
                          "target": {
                            "type": "RECTANGULAR_WINDOW",
                            "sheetName": "Budget",
                            "topLeftAddress": "A1",
                            "rowCount": 2.5,
                            "columnCount": 2
                          },
                          "query": { "type": "GET_WINDOW" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Unknown type value 'NO_SUCH_QUERY'", unknownType.getMessage());
    assertEquals("Field 'rowCount' must be an integer value", fractionalInteger.getMessage());
    assertEquals("steps[0].target.rowCount", fractionalInteger.jsonPath());
  }

  @Test
  void rejectsStepsThatMixActionAndQuery() {
    InvalidRequestShapeException invalidStep =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "bad",
                          "target": { "type": "CURRENT" },
                          "action": {
                            "type": "CLEAR_WORKBOOK_PROTECTION"
                          },
                          "query": { "type": "GET_WORKBOOK_SUMMARY" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals(
        "Each step must contain exactly one of 'action', 'assertion', or 'query'",
        invalidStep.getMessage());
    assertEquals("steps[0]", invalidStep.jsonPath());
  }

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
                NullPointerException.class, () -> GridGrindJson.writeProtocolCatalogBytes(null))
            .getMessage());
    assertEquals(
        "entry must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.writeTypeEntry(new ByteArrayOutputStream(), null))
            .getMessage());

    try (TrackingOutputStream outputStream = new TrackingOutputStream()) {
      GridGrindJson.writeRequest(
          outputStream,
          new WorkbookPlan(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.None(),
              List.of()));
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

    assertEquals(1, invalidResponse.jsonLine());
    assertEquals(2, invalidResponse.jsonColumn());
    org.junit.jupiter.api.Assertions.assertTrue(
        invalidCatalog.getMessage().startsWith("Missing required field '"));
  }

  @Test
  void exposesProductOwnedHelperMessagesAndJsonLocations() throws IOException {
    JsonFactory jsonFactory = new JsonFactory();
    MismatchedInputException nullPrimitive =
        (MismatchedInputException)
            MismatchedInputException.from(
                    jsonFactory.createParser("null"),
                    Integer.class,
                    "Cannot map `null` into type `int`")
                .prependPath(WorkbookPlan.class, "rowCount");
    MismatchedInputException floatingInteger =
        (MismatchedInputException)
            MismatchedInputException.from(
                    jsonFactory.createParser("2.5"),
                    Integer.class,
                    "Floating-point value (2.5) out of range of int")
                .prependPath(WorkbookPlan.class, "rowCount");
    InvalidTypeIdException invalidType =
        InvalidTypeIdException.from(jsonFactory.createParser("\"x\""), "bad type", null, "NOPE");

    assertEquals(
        "Missing required field 'rowCount'", GridGrindJson.mismatchedInputMessage(nullPrimitive));
    assertEquals(
        "Field 'rowCount' must be an integer value",
        GridGrindJson.mismatchedInputMessage(floatingInteger));
    assertEquals("Unknown type value 'NOPE'", GridGrindJson.message(invalidType));
    assertEquals(
        "Unknown field 'reads'",
        GridGrindJson.message(
            UnrecognizedPropertyException.from(
                jsonFactory.createParser("{}"), WorkbookPlan.class, "reads", List.of())));
    assertEquals(
        "JSON object is missing required fields or has the wrong shape",
        GridGrindJson.message(new IllegalArgumentException("Cannot construct instance of `x`")));
    assertEquals(
        "Cannot deserialize value",
        GridGrindJson.cleanJacksonMessage(
            "Cannot deserialize value as a subtype of `x` (for POJO property 'target') (set x.y to allow)"));
    assertEquals("Invalid JSON payload", GridGrindJson.cleanJacksonMessage(" "));
    assertNull(GridGrindJson.jsonLine(null));
    assertNull(GridGrindJson.jsonColumn(null));
    assertEquals(4, GridGrindJson.jsonLine(new TokenStreamLocation(null, 0L, 0L, 4, 9)));
    assertEquals(9, GridGrindJson.jsonColumn(new TokenStreamLocation(null, 0L, 0L, 4, 9)));
  }

  /** Tracks whether the request reader closes the source stream after consuming it. */
  @SuppressWarnings("NotJavadoc")
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

  /** Tracks whether the request writer closes the destination stream after producing JSON. */
  @SuppressWarnings("NotJavadoc")
  private static final class TrackingOutputStream extends ByteArrayOutputStream {
    private boolean closed;

    @Override
    public void close() {
      closed = true;
    }
  }
}
