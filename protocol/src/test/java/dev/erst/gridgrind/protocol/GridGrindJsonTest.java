package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
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
                    new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, true)),
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
        new BigDecimal("11.5"), applyStyle.style().fontHeight().toExcelFontHeight().points());
    assertEquals("#FFF2CC", applyStyle.style().fillColor());
    assertEquals(ExcelBorderStyle.THIN, applyStyle.style().border().all().style());
    assertEquals(ExcelBorderStyle.DOUBLE, applyStyle.style().border().right().style());
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

    assertNotNull(failure.getMessage());
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

    assertNotNull(failure.getMessage());
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

    assertNotNull(failure.getMessage());
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
                    GridGrindResponse.Problem.of(
                        GridGrindProblemCode.INVALID_REQUEST,
                        "bad",
                        new GridGrindResponse.ProblemContext.ValidateRequest(null, null)))));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeResponse(new ByteArrayOutputStream(), null));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeProtocolCatalog(null, GridGrindProtocolCatalog.catalog()));
    assertThrows(
        NullPointerException.class,
        () -> GridGrindJson.writeProtocolCatalog(new ByteArrayOutputStream(), null));
  }

  @Test
  void rejectsUnknownJsonFieldsInRequest() {
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
}
