package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
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
              "source": { "mode": "NEW" },
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
                      "source": { "mode": "EXISTING", "path": "budget.xlsm" },
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
              "source": { "mode": "NEW" },
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
              "source": { "mode": "NEW" },
              "operations": [],
              "reads": [
                { "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" },
                {
                  "type": "GET_NAMED_RANGES",
                  "requestId": "named-ranges",
                  "selection": {
                    "mode": "SELECTED",
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
                  "selection": { "mode": "ALL" }
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
  void wrapsRecordConstructionFailureAsInvalidPayloadError() {
    InvalidJsonException failure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readResponse(
                    """
                    {"status":"SUCCESS","protocolVersion":"V1","persistence":{"mode":"NOT_SAVED"},\
                    "reads":[{"type":"WORKBOOK_SUMMARY","requestId":"workbook","workbook":{"sheetCount":0,\
                    "sheetNames":[],"namedRangeCount":0,"forceFormulaRecalculationOnOpen":false}},\
                    {"type":"WINDOW","requestId":"window","window":{"sheetName":"Budget","topLeftAddress":"A1",\
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
  void wrapsDateTimeParseFailuresAsInvalidRequestErrors() {
    InvalidRequestException failure =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "mode": "NEW" },
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
  void stripsJacksonSourceLocationFromInvalidEnumValueMessages() {
    InvalidJsonException failure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "INVALID_VERSION",
                      "source": { "mode": "NEW" },
                      "operations": [],
                      "reads": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertNotNull(failure.getMessage());
    assertFalse(failure.getMessage().contains("[Source:"));
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
}
