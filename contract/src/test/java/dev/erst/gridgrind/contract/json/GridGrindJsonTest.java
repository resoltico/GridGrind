package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
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
import java.util.Optional;
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
  void readsRequestsFromAnExplicitTopLevelProtocolEnvelope() throws IOException {
    WorkbookPlan plan =
        GridGrindJson.readRequest(
            """
            {
              "protocolVersion": "V1",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
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
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(GridGrindProtocolVersion.V1, plan.protocolVersion());
  }

  @Test
  void rejectsRequestsThatOmitRequiredTopLevelSections() {
    InvalidRequestException exception =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" }
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("protocolVersion"));
  }

  @Test
  void rejectsNonObjectTopLevelPayloadsForWorkbookPlans() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJsonCodecSupport.readValue(
                    "[]".getBytes(StandardCharsets.UTF_8),
                    GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
                    WorkbookPlan.class,
                    GridGrindJsonMessageSupport::invalidRequestPayload));
    assertEquals("JSON value has the wrong shape for this field", exception.getMessage());
  }

  @Test
  void rejectsExplicitNullPlaceholdersDuringRequestRead() {
    InvalidRequestException topLevelNull =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "execution": null,
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestException nestedNull =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "summary",
                          "target": null,
                          "query": { "type": "GET_WORKBOOK_SUMMARY" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Missing required field 'execution'", topLevelNull.getMessage());
    assertEquals("Missing required field 'steps[0].target'", nestedNull.getMessage());
  }

  @Test
  @SuppressWarnings("StringConcatToTextBlock")
  void roundTripsRequestsResponsesAndCatalogs() throws IOException {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(
                new ExecutionModeInput(
                    ExecutionModeInput.ReadMode.FULL_XSSF, ExecutionModeInput.WriteMode.FULL_XSSF)),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-owner",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(new CellInput.Text(text("Owner")))),
                new AssertionStep(
                    "assert-owner",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));
    GridGrindResponse response =
        GridGrindResponses.success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponsePersistence.PersistenceOutcome.NotSaved(),
            List.of(new RequestWarning(0, "set-owner", "SET_CELL", "warning")),
            List.of(new AssertionResult("assert-owner", "EXPECT_CELL_VALUE")),
            List.of(
                new InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindWorkbookSurfaceReports.WorkbookSummary.WithSheets(
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
        GridGrindResponses.failure(
            GridGrindProtocolVersion.V1,
            new GridGrindProblemDetail.Problem(
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND
                    .category(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND
                    .recovery(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND.title(),
                "missing payload",
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND
                    .resolution(),
                new dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE"),
                    dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference
                        .path("cell text", "missing.txt")),
                java.util.Optional.empty(),
                List.of()));
    GridGrindResponse calculationFailure =
        GridGrindResponses.failure(
            GridGrindProtocolVersion.V1,
            new GridGrindProblemDetail.Problem(
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA,
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.category(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.recovery(),
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.title(),
                "bad formula",
                dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_FORMULA.resolution(),
                new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation.Preflight(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("EXISTING", "SAVE_AS"),
                    dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
                        .formulaCell("Ops", "B1", "SUM(")),
                java.util.Optional.empty(),
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
              "protocolVersion": "V1",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
                "journal": { "level": "NORMAL" },
                "calculation": {
                  "strategy": { "type": "EVALUATE_ALL" },
                  "markRecalculateOnOpen": false
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
              "protocolVersion": "V1",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
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
                      "source": { "type": "INLINE", "text": "Owner" }
                    }
                  }
                },
                {
                  "stepId": "assert-owner",
                  "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                  "assertion": {
                    "type": "EXPECT_CELL_VALUE",
                    "expectedValue": { "type": "TEXT", "text": "Owner" }
                  }
                },
                {
                  "stepId": "summary",
                  "target": { "type": "WORKBOOK_CURRENT" },
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
              "protocolVersion": "V1",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
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
                          "target": { "type": "WORKBOOK_CURRENT" },
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
                            "type": "RANGE_RECTANGULAR_WINDOW",
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
  void surfacesProductOwnedSuggestionsForCommonFirstContactTypeMistakes() {
    InvalidRequestShapeException legacyAssertion =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "legacy",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "assertion": { "type": "EXPECT_PRESENT" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException legacyAbsentAssertion =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "legacy-absent",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "assertion": { "type": "EXPECT_ABSENT" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException unknownAssertionType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "legacy-unknown",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "assertion": { "type": "EXPECT_LEGACYISH" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException wrongSourceType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": {
                        "type": "FILE",
                        "path": "budget.xlsx"
                      },
                      "steps": [ ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException unknownSourceType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": {
                        "type": "ARCHIVE",
                        "path": "budget.xlsx"
                      },
                      "steps": [ ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(
        legacyAssertion.getMessage().contains("EXPECT_CHART_PRESENT"),
        "legacy assertion failures must point at one of the new explicit family assertions");
    assertTrue(
        legacyAbsentAssertion.getMessage().contains("EXPECT_CHART_ABSENT"),
        "legacy absence failures must point at one of the new explicit family assertions");
    assertEquals("Unknown type value 'EXPECT_LEGACYISH'", unknownAssertionType.getMessage());
    assertTrue(
        wrongSourceType.getMessage().contains("source.type='EXISTING'"),
        "source type failures must teach the existing-workbook discriminator");
    assertEquals("Unknown type value 'ARCHIVE'", unknownSourceType.getMessage());
  }

  @Test
  void rejectsMissingPolymorphicAssertionTypeWithoutCrashing() {
    InvalidRequestShapeException missingAssertionType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "assertion-missing-type",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "assertion": { "F5pe": "EXPECT_CHART_PRESENT" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Missing required field 'type'", missingAssertionType.getMessage());
    assertEquals("steps[0].assertion", missingAssertionType.jsonPath());
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
                          "target": { "type": "WORKBOOK_CURRENT" },
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
  void rejectsDataValidationPayloadsThatOmitExplicitValidationBooleans() {
    InvalidRequestException missingAllowBlank =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "validation-missing-allowBlank",
                          "target": {
                            "type": "RANGE_BY_RANGE",
                            "sheetName": "Intake",
                            "range": "A2:A20"
                          },
                          "action": {
                            "type": "SET_DATA_VALIDATION",
                            "validation": {
                              "rule": {
                                "type": "EXPLICIT_LIST",
                                "values": [ "Queued" ]
                              },
                              "suppressDropDownArrow": false
                            }
                          }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestException missingSuppressDropDownArrow =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "validation-missing-suppressDropDownArrow",
                          "target": {
                            "type": "RANGE_BY_RANGE",
                            "sheetName": "Intake",
                            "range": "A2:A20"
                          },
                          "action": {
                            "type": "SET_DATA_VALIDATION",
                            "validation": {
                              "rule": {
                                "type": "EXPLICIT_LIST",
                                "values": [ "Queued" ]
                              },
                              "allowBlank": false
                            }
                          }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Missing required field 'allowBlank'", missingAllowBlank.getMessage());
    assertEquals(
        "steps[0].action.validation",
        missingAllowBlank.jsonPath(),
        "missing fields must point at the validation object");
    assertEquals(
        "Missing required field 'suppressDropDownArrow'",
        missingSuppressDropDownArrow.getMessage());
    assertEquals("steps[0].action.validation", missingSuppressDropDownArrow.jsonPath());
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
        "value must not be null",
        assertThrows(
                NullPointerException.class,
                () -> GridGrindJson.writeTypeEntry(new ByteArrayOutputStream(), null))
            .getMessage());

    try (TrackingOutputStream outputStream = new TrackingOutputStream()) {
      GridGrindJson.writeRequest(
          outputStream,
          WorkbookPlan.standard(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.None(),
              dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
              dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
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
    InvalidRequestException invalidCatalog =
        assertThrows(
            InvalidRequestException.class,
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
    assertEquals(Optional.empty(), GridGrindJson.jsonLine(null));
    assertEquals(Optional.empty(), GridGrindJson.jsonColumn(null));
    assertEquals(
        Optional.of(4), GridGrindJson.jsonLine(new TokenStreamLocation(null, 0L, 0L, 4, 9)));
    assertEquals(
        Optional.of(9), GridGrindJson.jsonColumn(new TokenStreamLocation(null, 0L, 0L, 4, 9)));
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
