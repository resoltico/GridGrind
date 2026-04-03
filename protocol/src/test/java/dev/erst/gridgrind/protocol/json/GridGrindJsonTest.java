package dev.erst.gridgrind.protocol.json;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.protocol.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.*;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.core.JacksonException;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.io.ContentReference;
import tools.jackson.databind.exc.MismatchedInputException;

/** Tests for GridGrindJson serialization and deserialization. */
class GridGrindJsonTest {
  @Test
  void defaultsRequestProtocolVersionDuringJsonRead() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "operations": [],
              "reads": []
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
              "source": { "type": "NEW" },
              "operations": [],
              "reads": []
            }
            """
                .getBytes(StandardCharsets.UTF_8))) {
      GridGrindRequest request = GridGrindJson.readRequest(inputStream);

      assertEquals(GridGrindProtocolVersion.V1, request.protocolVersion());
      assertFalse(inputStream.closed);
    }
  }

  @Test
  void readsProtocolCatalogsFromInputStreamsWithoutClosingThem() throws IOException {
    byte[] bytes = GridGrindJson.writeProtocolCatalogBytes(GridGrindProtocolCatalog.catalog());

    try (TrackingInputStream inputStream = new TrackingInputStream(bytes)) {
      GridGrindProtocolCatalog.Catalog catalog = GridGrindJson.readProtocolCatalog(inputStream);

      assertEquals(GridGrindProtocolCatalog.catalog(), catalog);
      assertFalse(inputStream.closed);
    }
  }

  @Test
  void roundTripsRequestTemplatesAndProtocolCatalogs() throws IOException {
    GridGrindRequest template = GridGrindProtocolCatalog.requestTemplate();
    GridGrindProtocolCatalog.Catalog catalog = GridGrindProtocolCatalog.catalog();

    GridGrindRequest decodedTemplate =
        GridGrindJson.readRequest(GridGrindJson.writeRequestBytes(template));
    GridGrindProtocolCatalog.Catalog decodedCatalog =
        GridGrindJson.readProtocolCatalog(GridGrindJson.writeProtocolCatalogBytes(catalog));

    assertEquals(template, decodedTemplate);
    assertEquals(catalog, decodedCatalog);
  }

  @Test
  void roundTripsStructuredResponses() throws IOException {
    GridGrindResponse expected =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs("budget.xlsx", "/tmp/budget.xlsx"),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "workbook",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, true)),
                new WorkbookReadResult.NamedRangesResult(
                    "named-ranges",
                    List.of(
                        new GridGrindResponse.NamedRangeReport.RangeReport(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            "Budget!$B$4",
                            new NamedRangeTarget("Budget", "B4"))))));

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
    InvalidJsonException responseFailure =
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readResponse("{".getBytes(StandardCharsets.UTF_8)));

    assertNotNull(requestFailure.getMessage());
    assertNotNull(responseFailure.getMessage());
    assertFalse(requestFailure.getMessage().contains("[Source:"));
    assertNull(requestFailure.jsonPath());
    assertEquals(1, requestFailure.jsonLine());
    assertEquals(2, requestFailure.jsonColumn());
  }

  @Test
  void wrapsMalformedJsonInputStreamsAsInvalidPayloadErrors() {
    InvalidJsonException requestFailure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readRequest(
                    new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8))));
    InvalidJsonException responseFailure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readResponse(
                    new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8))));

    assertNotNull(requestFailure.getMessage());
    assertNotNull(responseFailure.getMessage());
  }

  @Test
  void wrapsMalformedProtocolCatalogJsonAsInvalidPayloadErrors() {
    InvalidJsonException byteFailure =
        assertThrows(
            InvalidJsonException.class,
            () -> GridGrindJson.readProtocolCatalog("{".getBytes(StandardCharsets.UTF_8)));
    InvalidJsonException streamFailure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readProtocolCatalog(
                    new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8))));

    assertNotNull(byteFailure.getMessage());
    assertNotNull(streamFailure.getMessage());
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
                  "source": { "type": "NEW" },
                  "operations": [],
                  "reads": [
                    {
                      "type": "GET_WINDOW",
                      "requestId": "window",
                      "sheetName": "Budget",
                      "topLeftAddress": "A1",
                      "rowCount": 0,
                      "columnCount": 1
                    }
                  ]
                }
                """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("rowCount must be greater than 0", requestFailure.getMessage());
    assertEquals("reads[0]", requestFailure.jsonPath());
  }

  @Test
  void wrapsNonXlsxWorkbookPathsAsInvalidRequestErrors() {
    InvalidRequestException sourceFailure =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "EXISTING", "path": "budget.xlsm" },
                      "operations": [],
                      "reads": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    assertEquals(
        "path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: budget.xlsm",
        sourceFailure.getMessage());
    assertEquals("source", sourceFailure.jsonPath());
  }

  @Test
  void readsWorkbookOperationsFromJson() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "operations": [
                { "type": "RENAME_SHEET", "sheetName": "Budget", "newSheetName": "Summary" },
                { "type": "MERGE_CELLS", "sheetName": "Summary", "range": "A1:B2" },
                {
                  "type": "APPLY_STYLE",
                  "sheetName": "Summary",
                  "range": "A1",
                  "style": {
                    "fontName": "Aptos",
                    "fontHeight": { "type": "POINTS", "points": 11.5 },
                    "fillColor": "#fff2cc",
                    "border": {
                      "all": { "style": "THIN" },
                      "right": { "style": "DOUBLE" }
                    }
                  }
                },
                {
                  "type": "SET_HYPERLINK",
                  "sheetName": "Summary",
                  "address": "A1",
                  "target": { "type": "URL", "target": "https://example.com/report" }
                },
                {
                  "type": "SET_NAMED_RANGE",
                  "name": "BudgetTotal",
                  "scope": { "type": "WORKBOOK" },
                  "target": { "sheetName": "Summary", "range": "B4" }
                }
              ],
              "reads": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(5, request.operations().size());
    assertInstanceOf(WorkbookOperation.RenameSheet.class, request.operations().get(0));
    assertInstanceOf(WorkbookOperation.MergeCells.class, request.operations().get(1));
    assertInstanceOf(WorkbookOperation.ApplyStyle.class, request.operations().get(2));
    assertInstanceOf(WorkbookOperation.SetHyperlink.class, request.operations().get(3));
    assertInstanceOf(WorkbookOperation.SetNamedRange.class, request.operations().get(4));

    WorkbookOperation.ApplyStyle applyStyle =
        (WorkbookOperation.ApplyStyle) request.operations().get(2);
    assertEquals("Aptos", applyStyle.style().fontName());
    assertEquals(
        new BigDecimal("11.5"),
        assertInstanceOf(FontHeightInput.Points.class, applyStyle.style().fontHeight()).points());
    assertEquals("#FFF2CC", applyStyle.style().fillColor());
    assertEquals(BorderStyle.THIN, applyStyle.style().border().all().style());
    assertEquals(BorderStyle.DOUBLE, applyStyle.style().border().right().style());
  }

  @Test
  void readsSheetStateOperationsFromJson() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "operations": [
                {
                  "type": "COPY_SHEET",
                  "sourceSheetName": "Budget",
                  "newSheetName": "Budget Copy",
                  "position": { "type": "AT_INDEX", "targetIndex": 1 }
                },
                { "type": "SET_ACTIVE_SHEET", "sheetName": "Budget Copy" },
                { "type": "SET_SELECTED_SHEETS", "sheetNames": ["Budget", "Budget Copy"] },
                {
                  "type": "SET_SHEET_VISIBILITY",
                  "sheetName": "Budget",
                  "visibility": "VERY_HIDDEN"
                },
                {
                  "type": "SET_SHEET_PROTECTION",
                  "sheetName": "Budget",
                  "protection": {
                    "autoFilterLocked": false,
                    "deleteColumnsLocked": true,
                    "deleteRowsLocked": false,
                    "formatCellsLocked": true,
                    "formatColumnsLocked": false,
                    "formatRowsLocked": true,
                    "insertColumnsLocked": false,
                    "insertHyperlinksLocked": true,
                    "insertRowsLocked": false,
                    "objectsLocked": true,
                    "pivotTablesLocked": false,
                    "scenariosLocked": true,
                    "selectLockedCellsLocked": false,
                    "selectUnlockedCellsLocked": true,
                    "sortLocked": false
                  }
                },
                { "type": "CLEAR_SHEET_PROTECTION", "sheetName": "Budget" }
              ],
              "reads": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(6, request.operations().size());
    WorkbookOperation.CopySheet copySheet =
        assertInstanceOf(WorkbookOperation.CopySheet.class, request.operations().get(0));
    WorkbookOperation.SetActiveSheet setActiveSheet =
        assertInstanceOf(WorkbookOperation.SetActiveSheet.class, request.operations().get(1));
    WorkbookOperation.SetSelectedSheets setSelectedSheets =
        assertInstanceOf(WorkbookOperation.SetSelectedSheets.class, request.operations().get(2));
    WorkbookOperation.SetSheetVisibility setSheetVisibility =
        assertInstanceOf(WorkbookOperation.SetSheetVisibility.class, request.operations().get(3));
    WorkbookOperation.SetSheetProtection setSheetProtection =
        assertInstanceOf(WorkbookOperation.SetSheetProtection.class, request.operations().get(4));
    WorkbookOperation.ClearSheetProtection clearSheetProtection =
        assertInstanceOf(WorkbookOperation.ClearSheetProtection.class, request.operations().get(5));

    SheetCopyPosition.AtIndex position =
        assertInstanceOf(SheetCopyPosition.AtIndex.class, copySheet.position());
    assertEquals(1, position.targetIndex());
    assertEquals("Budget Copy", setActiveSheet.sheetName());
    assertEquals(List.of("Budget", "Budget Copy"), setSelectedSheets.sheetNames());
    assertEquals(SheetVisibility.VERY_HIDDEN, setSheetVisibility.visibility());
    assertEquals(protectionSettings(), setSheetProtection.protection());
    assertEquals("Budget", clearSheetProtection.sheetName());
  }

  @Test
  void readsWorkbookReadOperationsFromJson() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "operations": [],
              "reads": [
                { "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" },
                {
                  "type": "GET_NAMED_RANGES",
                  "requestId": "named-ranges",
                  "selection": {
                    "type": "SELECTED",
                    "selectors": [
                      { "type": "WORKBOOK_SCOPE", "name": "BudgetTotal" }
                    ]
                  }
                },
                {
                  "type": "GET_WINDOW",
                  "requestId": "window",
                  "sheetName": "Budget",
                  "topLeftAddress": "A1",
                  "rowCount": 4,
                  "columnCount": 3
                },
                {
                  "type": "GET_FORMULA_SURFACE",
                  "requestId": "formula-surface",
                  "selection": { "type": "ALL" }
                },
                {
                  "type": "ANALYZE_WORKBOOK_FINDINGS",
                  "requestId": "workbook-findings"
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(5, request.reads().size());
    assertInstanceOf(WorkbookReadOperation.GetWorkbookSummary.class, request.reads().get(0));
    assertInstanceOf(WorkbookReadOperation.GetNamedRanges.class, request.reads().get(1));
    assertInstanceOf(WorkbookReadOperation.GetWindow.class, request.reads().get(2));
    assertInstanceOf(WorkbookReadOperation.GetFormulaSurface.class, request.reads().get(3));
    assertInstanceOf(WorkbookReadOperation.AnalyzeWorkbookFindings.class, request.reads().get(4));
  }

  @Test
  void readsFileHyperlinksFromPathFieldsAndNormalizesFileUris() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "operations": [
                {
                  "type": "SET_HYPERLINK",
                  "sheetName": "Budget",
                  "address": "A1",
                  "target": { "type": "FILE", "path": "file:///tmp/report.xlsx" }
                }
              ],
              "reads": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    WorkbookOperation.SetHyperlink operation =
        assertInstanceOf(WorkbookOperation.SetHyperlink.class, request.operations().getFirst());
    assertEquals(new HyperlinkTarget.File("/tmp/report.xlsx"), operation.target());
  }

  @Test
  void writesFileHyperlinksToResponsesWithPathFields() throws IOException {
    GridGrindResponse response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            List.of(
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.File("/tmp/report.xlsx"))))));

    byte[] encoded = GridGrindJson.writeResponseBytes(response);
    String json = new String(encoded, StandardCharsets.UTF_8);
    GridGrindResponse decoded = GridGrindJson.readResponse(encoded);
    WorkbookReadResult.HyperlinksResult hyperlinks =
        assertInstanceOf(
            WorkbookReadResult.HyperlinksResult.class,
            assertInstanceOf(GridGrindResponse.Success.class, decoded).reads().getFirst());

    assertTrue(json.contains("\"path\" : \"/tmp/report.xlsx\""));
    assertFalse(json.contains("\"target\" : \"/tmp/report.xlsx\""));
    assertEquals(
        new HyperlinkTarget.File("/tmp/report.xlsx"),
        hyperlinks.hyperlinks().getFirst().hyperlink());
  }

  @Test
  void writesProblemCausesUsingGridGrindCodesInsteadOfExceptionMetadata() throws IOException {
    GridGrindResponse response =
        new GridGrindResponse.Failure(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.Problem(
                GridGrindProblemCode.INVALID_REQUEST,
                GridGrindProblemCategory.REQUEST,
                GridGrindProblemRecovery.CHANGE_REQUEST,
                "Invalid request",
                "bad request",
                "Fix the request.",
                new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"),
                List.of(
                    new GridGrindResponse.ProblemCause(
                        GridGrindProblemCode.INTERNAL_ERROR,
                        "Internal GridGrind failure",
                        "EXECUTE_REQUEST"))));

    String json = new String(GridGrindJson.writeResponseBytes(response), StandardCharsets.UTF_8);
    GridGrindResponse decoded = GridGrindJson.readResponse(json.getBytes(StandardCharsets.UTF_8));
    GridGrindResponse.Failure failure = assertInstanceOf(GridGrindResponse.Failure.class, decoded);

    assertTrue(json.contains("\"code\" : \"INTERNAL_ERROR\""));
    assertFalse(json.contains("\"className\""));
    assertFalse(json.contains("\"type\" : \"GridGrindProblem\""));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().causes().getFirst().code());
  }

  @Test
  void wrapsResponseShapeFailuresAsInvalidRequestShapeErrors() {
    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readResponse(
                    """
                    {"status":"SUCCESS","protocolVersion":"V1","persistence":{"type":"NONE"},\
                    "reads":[{"type":"GET_WORKBOOK_SUMMARY","requestId":"workbook","workbook":{"sheetCount":0,\
                    "sheetNames":[],"namedRangeCount":0,"forceFormulaRecalculationOnOpen":false}},\
                    {"type":"GET_WINDOW","requestId":"window","window":{"sheetName":"Budget","topLeftAddress":"A1",\
                    "rowCount":1,"columnCount":1,"rows":[{"rowIndex":0,"cells":[{"effectiveType":"BLANK",\
                    "address":"A1","declaredType":"BLANK","displayValue":"","style":{"numberFormat":null,\
                    "bold":false,"italic":false,"wrapText":false,"horizontalAlignment":"GENERAL",\
                    "verticalAlignment":"BOTTOM","fontName":"Calibri","fontHeight":{"twips":220,"points":11},\
                    "fontColor":null,"underline":false,"strikeout":false,"fillColor":null,\
                    "topBorderStyle":"NONE","rightBorderStyle":"NONE","bottomBorderStyle":"NONE",\
                    "leftBorderStyle":"NONE"}}]}]}}]}"""
                        .getBytes(StandardCharsets.UTF_8)));

    assertNotNull(failure.getMessage());
    assertFalse(failure.getMessage().contains("[Source:"));
  }

  @Test
  void wrapsProtocolCatalogShapeFailuresAsInvalidRequestShapeErrors() {
    InvalidRequestShapeException byteFailure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readProtocolCatalog(
                    """
                    {
                      "protocolVersion": "V1",
                      "sourceTypes": [],
                      "persistenceTypes": [],
                      "operationTypes": [],
                      "readTypes": [],
                      "nestedTypes": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    InvalidRequestShapeException streamFailure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readProtocolCatalog(
                    new ByteArrayInputStream(
                        """
                        {
                          "protocolVersion": "V1",
                          "sourceTypes": [],
                          "persistenceTypes": [],
                          "operationTypes": [],
                          "readTypes": [],
                          "nestedTypes": []
                        }
                        """
                            .getBytes(StandardCharsets.UTF_8))));

    assertNull(byteFailure.jsonPath());
    assertNull(streamFailure.jsonPath());
  }

  @Test
  void wrapsUnknownDiscriminatorsAsInvalidRequestShapeErrors() {
    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "operations": [],
                      "reads": [
                        { "type": "GET_NOT_A_REAL_READ", "requestId": "workbook" }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Unknown type value 'GET_NOT_A_REAL_READ'", failure.getMessage());
    assertFalse(failure.getMessage().contains("dev.erst.gridgrind"));
    assertFalse(failure.getMessage().contains("for POJO property"));
    assertEquals("reads[0]", failure.jsonPath());
  }

  @Test
  void wrapsMissingRequiredFieldsAsInvalidRequestShapeErrors() {
    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "operations": [],
                      "reads": [
                        { "type": "GET_WORKBOOK_SUMMARY" }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("Missing required field 'requestId'", failure.getMessage());
    assertFalse(failure.getMessage().contains("tools.jackson"));
    assertFalse(failure.getMessage().contains("dev.erst.gridgrind"));
    assertEquals("reads[0]", failure.jsonPath());
  }

  @Test
  void wrapsWrongTokenShapesAsInvalidRequestShapeErrors() {
    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "operations": [
                        {
                          "type": "SET_CELL",
                          "sheetName": { "value": "Budget" },
                          "address": "A1",
                          "value": { "type": "TEXT", "text": "Report" }
                        }
                      ],
                      "reads": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertEquals("JSON value has the wrong shape for this field", failure.getMessage());
    assertEquals("operations[0].sheetName", failure.jsonPath());
  }

  @Test
  void wrapsDateTimeParseFailuresAsInvalidRequestErrors() {
    InvalidRequestException failure =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "operations": [
                        {
                          "type": "SET_CELL",
                          "sheetName": "Budget",
                          "address": "A1",
                          "value": { "type": "DATE", "date": "not-a-date" }
                        }
                      ],
                      "reads": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertNotNull(failure.getMessage());
  }

  @Test
  void stripsJacksonSourceLocationFromInvalidRequestShapeMessages() {
    InvalidRequestShapeException failure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "INVALID_VERSION",
                      "source": { "type": "NEW" },
                      "operations": [],
                      "reads": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertNotNull(failure.getMessage());
    assertFalse(failure.getMessage().contains("[Source:"));
  }

  @Test
  void wrapsUnexpectedJacksonExceptionsAsInvalidJsonErrors() {
    InvalidJsonException failure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readRequest(
                    new ThrowingJacksonInputStream(unexpectedJacksonException())));

    assertEquals("synthetic failure", failure.getMessage());
    assertEquals("source", failure.jsonPath());
    assertEquals(7, failure.jsonLine());
    assertEquals(9, failure.jsonColumn());
  }

  @Test
  void writesRequestsAndProtocolCatalogsToStreamsWithoutClosingThem() throws IOException {
    try (TrackingOutputStream requestOutput = new TrackingOutputStream();
        TrackingOutputStream catalogOutput = new TrackingOutputStream()) {
      GridGrindJson.writeRequest(requestOutput, GridGrindProtocolCatalog.requestTemplate());
      GridGrindJson.writeProtocolCatalog(catalogOutput, GridGrindProtocolCatalog.catalog());

      assertEquals(
          GridGrindProtocolCatalog.requestTemplate(),
          GridGrindJson.readRequest(requestOutput.toByteArray()));
      assertEquals(
          GridGrindProtocolCatalog.catalog(),
          GridGrindJson.readProtocolCatalog(catalogOutput.toByteArray()));
      assertFalse(requestOutput.closed);
      assertFalse(catalogOutput.closed);
    }
  }

  @Test
  void roundTripsSheetStateReadResults() throws IOException {
    GridGrindResponse expected =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "workbook",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        2,
                        List.of("Budget", "Archive"),
                        "Archive",
                        List.of("Budget", "Archive"),
                        1,
                        true)),
                new WorkbookReadResult.SheetSummaryResult(
                    "sheet",
                    new GridGrindResponse.SheetSummaryReport(
                        "Budget",
                        SheetVisibility.VERY_HIDDEN,
                        new GridGrindResponse.SheetProtectionReport.Protected(protectionSettings()),
                        4,
                        7,
                        2))));

    byte[] encoded = GridGrindJson.writeResponseBytes(expected);
    String json = new String(encoded, StandardCharsets.UTF_8);
    GridGrindResponse.Success actual =
        assertInstanceOf(GridGrindResponse.Success.class, GridGrindJson.readResponse(encoded));
    WorkbookReadResult.WorkbookSummaryResult workbook =
        assertInstanceOf(WorkbookReadResult.WorkbookSummaryResult.class, actual.reads().get(0));
    GridGrindResponse.WorkbookSummary.WithSheets workbookSummary =
        assertInstanceOf(GridGrindResponse.WorkbookSummary.WithSheets.class, workbook.workbook());
    WorkbookReadResult.SheetSummaryResult sheet =
        assertInstanceOf(WorkbookReadResult.SheetSummaryResult.class, actual.reads().get(1));
    GridGrindResponse.SheetProtectionReport.Protected protection =
        assertInstanceOf(
            GridGrindResponse.SheetProtectionReport.Protected.class, sheet.sheet().protection());

    assertTrue(json.contains("\"type\" : \"GET_SHEET_SUMMARY\""));
    assertTrue(json.contains("\"kind\" : \"WITH_SHEETS\""));
    assertEquals("Archive", workbookSummary.activeSheetName());
    assertEquals(List.of("Budget", "Archive"), workbookSummary.selectedSheetNames());
    assertEquals(SheetVisibility.VERY_HIDDEN, sheet.sheet().visibility());
    assertEquals(protectionSettings(), protection.settings());
  }

  @Test
  void jsonLineAndColumnReturnNullForNullOrNonPositiveLocation() {
    assertNull(GridGrindJson.jsonLine(null));
    assertNull(GridGrindJson.jsonColumn(null));
    assertNull(GridGrindJson.jsonLine(tools.jackson.core.TokenStreamLocation.NA));
    assertNull(GridGrindJson.jsonColumn(tools.jackson.core.TokenStreamLocation.NA));
  }

  @Test
  void validatesNullArguments() {
    assertThrows(NullPointerException.class, () -> GridGrindJson.readRequest((byte[]) null));
    assertThrows(
        NullPointerException.class, () -> GridGrindJson.readRequest((ByteArrayInputStream) null));
    assertThrows(NullPointerException.class, () -> GridGrindJson.readResponse((byte[]) null));
    assertThrows(
        NullPointerException.class, () -> GridGrindJson.readResponse((ByteArrayInputStream) null));
    assertThrows(
        NullPointerException.class, () -> GridGrindJson.readProtocolCatalog((byte[]) null));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.readProtocolCatalog((ByteArrayInputStream) null));
    assertThrows(NullPointerException.class, () -> GridGrindJson.writeRequestBytes(null));
    assertThrows(NullPointerException.class, () -> GridGrindJson.writeResponseBytes(null));
    assertThrows(NullPointerException.class, () -> GridGrindJson.writeProtocolCatalogBytes(null));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeRequest(null, GridGrindProtocolCatalog.requestTemplate()));
    assertThrows(
        NullPointerException.class,
        () ->
            GridGrindJson.writeResponse(
                null,
                new GridGrindResponse.Failure(
                    null,
                    new GridGrindResponse.Problem(
                        GridGrindProblemCode.INVALID_REQUEST,
                        GridGrindProblemCategory.REQUEST,
                        GridGrindProblemRecovery.CHANGE_REQUEST,
                        "Invalid request",
                        "bad",
                        "Fix the request.",
                        new GridGrindResponse.ProblemContext.ValidateRequest(null, null),
                        List.of()))));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeResponse(new ByteArrayOutputStream(), null));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeProtocolCatalog(null, GridGrindProtocolCatalog.catalog()));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeProtocolCatalog(new ByteArrayOutputStream(), null));
    assertThrows(
        NullPointerException.class,
        () ->
            GridGrindJson.writeTypeEntry(
                null, GridGrindProtocolCatalog.catalog().operationTypes().getFirst()));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeTypeEntry(new ByteArrayOutputStream(), null));
  }

  @Test
  void writeTypeEntrySerializesASingleCatalogEntry() throws Exception {
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    GridGrindProtocolCatalog.TypeEntry entry =
        GridGrindProtocolCatalog.entryFor("SET_CELL").orElseThrow();
    GridGrindJson.writeTypeEntry(out, entry);
    String json = out.toString(java.nio.charset.StandardCharsets.UTF_8);
    assertTrue(json.contains("\"SET_CELL\""), "output must contain the entry id");
    assertTrue(json.contains("\"summary\""), "output must contain the summary field");
    assertTrue(json.contains("\"fields\""), "output must contain field descriptors");
  }

  @Test
  void rejectsUnknownJsonFieldsInRequest() {
    InvalidRequestShapeException topLevelFailure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "operations": [],
                      "reads": [],
                      "unknownTopLevelField": "surprise"
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    assertEquals("Unknown field 'unknownTopLevelField'", topLevelFailure.getMessage());

    InvalidRequestShapeException operationFailure =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "type": "NEW" },
                      "operations": [
                        { "type": "SET_CELL", "sheetName": "Budget", "address": "A1",
                          "value": { "type": "TEXT", "text": "hi" },
                          "dataFormat": "#,##0" }
                      ],
                      "reads": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    assertEquals("Unknown field 'dataFormat'", operationFailure.getMessage());
  }

  @Test
  void cleanJacksonMessageFallsBackForBlankMessages() {
    assertEquals("Invalid JSON payload", GridGrindJson.cleanJacksonMessage(null));
    assertEquals("Invalid JSON payload", GridGrindJson.cleanJacksonMessage(" "));
  }

  @Test
  void mismatchedInputMessageReturnsTheCleanedOriginalWhenItIsAlreadySpecific() {
    MismatchedInputException exception =
        MismatchedInputException.from(null, Object.class, "Expected an array value");

    assertEquals("Expected an array value", GridGrindJson.mismatchedInputMessage(exception));
  }

  @Test
  void mismatchedInputMessageMapsMissingCreatorPropertiesToMissingFieldMessages() {
    MismatchedInputException exception =
        MismatchedInputException.from(
            null, Object.class, "Missing required creator property 'requestId' (index 0)");

    assertEquals(
        "Missing required field 'requestId'", GridGrindJson.mismatchedInputMessage(exception));
  }

  @Test
  void mismatchedInputMessageMapsMissingTypeIdsToMissingFieldMessages() {
    MismatchedInputException exception =
        MismatchedInputException.from(null, Object.class, "Missing type id property 'type'");

    assertEquals("Missing required field 'type'", GridGrindJson.mismatchedInputMessage(exception));
  }

  @Test
  void mismatchedInputMessageMapsGenericConstructionFailuresToShapeMessages() {
    MismatchedInputException exception =
        MismatchedInputException.from(
            null,
            Object.class,
            "Cannot construct instance of `dev.erst.gridgrind.protocol.Shape`, problem: bad token");

    assertEquals(
        "JSON object is missing required fields or has the wrong shape",
        GridGrindJson.mismatchedInputMessage(exception));
  }

  /** ByteArrayInputStream that records whether {@code close()} was called. */
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

  /** ByteArrayOutputStream that records whether {@code close()} was called. */
  private static final class TrackingOutputStream extends ByteArrayOutputStream {
    private boolean closed;

    @Override
    public void close() throws IOException {
      closed = true;
      super.close();
    }
  }

  /** InputStream that throws a caller-supplied JacksonException on first read. */
  private static final class ThrowingJacksonInputStream extends InputStream {
    private final JacksonException exception;

    private ThrowingJacksonInputStream(JacksonException exception) {
      this.exception = exception;
    }

    @Override
    public int read() {
      throw exception;
    }

    @Override
    public int read(byte[] destination, int offset, int length) {
      throw exception;
    }
  }

  private static JacksonException unexpectedJacksonException() {
    SyntheticJacksonException exception =
        new SyntheticJacksonException(
            "synthetic failure",
            new TokenStreamLocation(ContentReference.rawReference("request.json"), 13L, 7, 9));
    exception.prependPath(GridGrindJsonTest.class, "source");
    return exception;
  }

  /** Synthetic Jackson exception used to exercise the fallback classifier path. */
  private static final class SyntheticJacksonException extends JacksonException {
    private static final long serialVersionUID = 1L;

    private SyntheticJacksonException(String message, TokenStreamLocation location) {
      super(message, location, null);
    }
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }
}
