package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for JSON serialization, parser wording, and the step-based wire shape. */
@SuppressWarnings("NotJavadoc")
class GridGrindJsonTest {
  @Test
  void readsRequestsFromTheMinimalTopLevelProtocolEnvelope() throws IOException {
    WorkbookPlan plan =
        GridGrindJson.readRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(GridGrindProtocolVersion.V2, plan.protocolVersion());
    assertTrue(plan.execution().isDefault());
    assertEquals(ExecutionJournalLevel.SUMMARY, plan.journalLevel());
    assertTrue(plan.formulaEnvironment().isEmpty());
    String serialized =
        new String(GridGrindJsonOutput.writeRequestBytes(plan), StandardCharsets.UTF_8);
    assertFalse(serialized.contains("\"execution\""));
    assertFalse(serialized.contains("\"formulaEnvironment\""));
  }

  @Test
  void rejectsRequestsThatOmitRequiredTopLevelSections() {
    InvalidRequestShapeException exception =
        assertThrows(
            InvalidRequestShapeException.class,
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
  void reportsUnsupportedEnumValuesWithAllowedCandidates() {
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
                        "mode": { "type": "FULL_XSSF" },
                        "journal": { "level": "SUMMARY" },
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
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals(
        "Unsupported value 'V1' for field 'protocolVersion'; expected one of: V2",
        exception.getMessage());
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
                    GridGrindJsonProblemMessageSupport::invalidRequestPayload));
    assertEquals("JSON value has the wrong shape for this field", exception.getMessage());
  }

  @Test
  void rejectsExplicitNullPlaceholdersDuringRequestRead() {
    InvalidRequestShapeException optionalNull =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "planId": null,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": { "type": "FULL_XSSF" },
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
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException topLevelNull =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": null,
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException nestedNull =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
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

    assertEquals(
        "Field 'planId' must be omitted when absent; explicit null is not accepted.",
        optionalNull.getMessage());
    assertEquals(
        "Field 'execution' must be omitted when absent; explicit null is not accepted.",
        topLevelNull.getMessage());
    assertEquals(
        "Field 'steps[0].target' must be omitted when absent; explicit null is not accepted.",
        nestedNull.getMessage());
    assertEquals(Optional.of("planId"), optionalNull.jsonPath());
    assertEquals(Optional.of("execution"), topLevelNull.jsonPath());
    assertEquals(Optional.of("steps[0].target"), nestedNull.jsonPath());
  }

  @Test
  void derivesJsonPathForMissingCreatorProperties() {
    IllegalArgumentException missingProtocolVersion =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": { "type": "FULL_XSSF" },
                        "journal": { "level": "SUMMARY" },
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
                        .getBytes(StandardCharsets.UTF_8)));

    assertInstanceOf(InvalidRequestShapeException.class, missingProtocolVersion);
    assertInstanceOf(PayloadException.class, missingProtocolVersion);
    PayloadException payloadException = (PayloadException) missingProtocolVersion;
    assertEquals("Missing required field 'protocolVersion'", missingProtocolVersion.getMessage());
    assertEquals(Optional.of("protocolVersion"), payloadException.jsonPath());
  }

  @Test
  void derivesJsonPathForMissingSaveAsIfExists() {
    IllegalArgumentException missingIfExists =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "SAVE_AS", "path": "budget.xlsx" },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertInstanceOf(InvalidRequestShapeException.class, missingIfExists);
    assertInstanceOf(PayloadException.class, missingIfExists);
    PayloadException payloadException = (PayloadException) missingIfExists;
    assertEquals("Missing required field 'persistence.ifExists'", missingIfExists.getMessage());
    assertEquals(Optional.of("persistence.ifExists"), payloadException.jsonPath());
  }

  @Test
  void derivesPreciseJsonPathForNonXlsxWorkbookPathViolations() {
    InvalidRequestException invalidSourcePath =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "EXISTING", "path": "budget.txt" },
                      "persistence": { "type": "NONE" },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestException invalidPersistencePath =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "SAVE_AS", "path": "budget.txt", "ifExists": "REJECT" },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestException invalidExternalWorkbookPath =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "formulaEnvironment": {
                        "externalWorkbooks": [
                          { "workbookName": "Rates.xlsx", "path": "rates.txt" }
                        ],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals(Optional.of("source.path"), invalidSourcePath.jsonPath());
    assertEquals(Optional.of("persistence.path"), invalidPersistencePath.jsonPath());
    assertEquals(
        Optional.of("formulaEnvironment.externalWorkbooks[0].path"),
        invalidExternalWorkbookPath.jsonPath(),
        invalidExternalWorkbookPath::toString);
  }

  @Test
  @SuppressWarnings("StringConcatToTextBlock")
  void roundTripsRequestsResponsesAndCatalogs() throws IOException {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(ExecutionModeInput.fullXssf()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-owner",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(new CellInput.Text(text("Owner")))),
                new AssertionStep(
                    "assert-owner",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellAssertion.CellValue(
                        new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"))),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    WorkbookResult response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
            List.of(
                new RequestWarning(
                    dev.erst.gridgrind.contract.dto.GridGrindWarningCode
                        .UNQUOTED_SHEET_NAME_IN_FORMULA,
                    0,
                    "set-owner",
                    "SET_CELL",
                    "warning")),
            List.of(new AssertionResult.Passed("assert-owner", "EXPECT_CELL_VALUE")),
            List.of(
                new WorkbookInspectionResult.WorkbookSummaryResult(
                    "summary",
                    new WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 0, false))));
    Catalog catalog = GridGrindProtocolCatalog.catalog();

    assertEquals(
        request, GridGrindJson.readRequest(GridGrindJsonOutput.writeRequestBytes(request)));
    assertEquals(
        response,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(response)));
    assertEquals(
        catalog,
        GridGrindJson.readProtocolCatalog(GridGrindJsonOutput.writeProtocolCatalogBytes(catalog)));
  }

  @Test
  void roundTripsResolveInputsAndCalculationFailureContexts() throws IOException {
    WorkbookResult resolveInputsFailure =
        WorkbookResults.failure(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
                "out/report.xlsx", new WorkbookResultPersistence.WriteResult.NotWritten()),
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
                        .path("cell text", "missing.txt"),
                    java.util.Optional.of(
                        dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                            .pathAtByteOffset("steps[0].action.value.source.path", 73))),
                List.of()));
    WorkbookResult calculationFailure =
        WorkbookResults.failure(
            GridGrindProtocolVersion.V2,
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
                List.of()));

    assertEquals(
        resolveInputsFailure,
        GridGrindJson.readWorkbookResult(
            GridGrindJsonOutput.writeWorkbookResultBytes(resolveInputsFailure)));
    assertEquals(
        calculationFailure,
        GridGrindJson.readWorkbookResult(
            GridGrindJsonOutput.writeWorkbookResultBytes(calculationFailure)));
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
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "mode": {"type": "FULL_XSSF"},
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

    byte[] serialized = GridGrindJsonOutput.writeRequestBytes(request);
    String serializedJson = new String(serialized, StandardCharsets.UTF_8);

    assertFalse(serializedJson.contains("\"default\""));
    assertEquals(ExecutionModeInput.fullXssf(), request.effectiveExecutionMode());
    assertFalse(request.calculationPolicy().markRecalculateOnOpen());
    assertEquals(request, GridGrindJson.readRequest(serialized));
  }

  @Test
  void readsCanonicalStepEnvelopeWithoutOuterStepType() throws IOException {
    WorkbookPlan plan =
        GridGrindJson.readRequest(
            """
            {
              "protocolVersion": "V2",
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
              "protocolVersion": "V2",
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
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8))) {
      WorkbookPlan request = GridGrindJson.readRequest(inputStream);

      assertEquals(GridGrindProtocolVersion.V2, request.protocolVersion());
      assertFalse(inputStream.closed);
    }
  }

  @Test
  void rejectsUnknownTypeValuesNonStringTypeIdsAndFractionalIntegersWithProductOwnedMessages() {
    InvalidRequestShapeException unknownType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
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
    InvalidRequestShapeException nonStringType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": [
                        {
                          "stepId": "bad-shape",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "query": { "type": 1 }
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
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
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
    assertEquals(Optional.of("steps[0].query.type"), unknownType.jsonPath());
    assertEquals(
        "Field 'steps[0].query.type' must be a JSON string type id", nonStringType.getMessage());
    assertEquals(Optional.of("steps[0].query.type"), nonStringType.jsonPath());
    assertEquals(
        "Field 'steps[0].target.rowCount' must be a JSON integer between -2147483648 and 2147483647",
        fractionalInteger.getMessage());
    assertEquals(Optional.of("steps[0].target.rowCount"), fractionalInteger.jsonPath());
  }

  @Test
  void surfacesProductOwnedSuggestionsForCommonFirstContactTypeMistakes() {
    InvalidRequestShapeException deletedAssertionName =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
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
    InvalidRequestShapeException deletedAbsentAssertionName =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
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
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
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
                      "protocolVersion": "V2",
                      "source": {
                        "type": "FILE",
                        "path": "budget.xlsx"
                      },
                      "persistence": { "type": "NONE" },
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
                      "protocolVersion": "V2",
                      "source": {
                        "type": "ARCHIVE",
                        "path": "budget.xlsx"
                      },
                      "persistence": { "type": "NONE" },
                      "steps": [ ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Unknown type value 'EXPECT_PRESENT'", deletedAssertionName.getMessage());
    assertEquals("Unknown type value 'EXPECT_ABSENT'", deletedAbsentAssertionName.getMessage());
    assertEquals("Unknown type value 'EXPECT_LEGACYISH'", unknownAssertionType.getMessage());
    assertTrue(
        wrongSourceType.getMessage().contains("source.type='EXISTING'"),
        "source type failures must teach the existing-workbook discriminator");
    assertEquals("Unknown type value 'ARCHIVE'", unknownSourceType.getMessage());
  }

  @Test
  void reportsMissingPolymorphicAssertionTypeAndUnknownMembersWithoutCrashing() {
    byte[] request =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "assertion-missing-type",
              "target": { "type": "WORKBOOK_CURRENT" },
              "assertion": { "F5pe": "EXPECT_CHART_PRESENT" }
            }
          ]
        }
        """
            .getBytes(StandardCharsets.UTF_8);

    RequestAnalysis analysis = GridGrindJson.analyzeRequest(request);
    assertEquals(
        List.of("steps[0].assertion.F5pe", "steps[0].assertion.type"),
        analysis.structuralProblems().stream()
            .map(problem -> problem.jsonPath().orElseThrow())
            .toList());
    assertEquals(
        List.of(RequestUnknownField.class, RequestMissingTypeDiscriminator.class),
        analysis.structuralProblems().stream().map(Object::getClass).toList());

    InvalidRequestShapeException firstProblem =
        assertThrows(InvalidRequestShapeException.class, () -> GridGrindJson.readRequest(request));
    assertEquals("Unknown field 'steps[0].assertion.F5pe'", firstProblem.getMessage());
    assertEquals(Optional.of("steps[0].assertion.F5pe"), firstProblem.jsonPath());
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
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
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
        "Field 'steps[0]' must be an object containing exactly one of action, assertion, or query",
        invalidStep.getMessage());
    assertEquals(Optional.of("steps[0]"), invalidStep.jsonPath());
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
}
