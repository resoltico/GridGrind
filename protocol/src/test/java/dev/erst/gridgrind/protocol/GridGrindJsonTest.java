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
            new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, true),
            List.of(
                new GridGrindResponse.NamedRangeReport.RangeReport(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    "Budget!$B$4",
                    new NamedRangeTarget("Budget", "B4"))),
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
                      "analysis": { "sheets": [] }
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    assertEquals(
        "path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: budget.xlsm",
        sourceFailure.getMessage());
    assertEquals("source", sourceFailure.jsonPath());

    InvalidRequestException persistenceFailure =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "source": { "mode": "NEW" },
                      "persistence": { "mode": "SAVE_AS", "path": "output.xls" },
                      "operations": [],
                      "analysis": { "sheets": [] }
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));
    assertEquals(
        "path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: output.xls",
        persistenceFailure.getMessage());
    assertEquals("persistence", persistenceFailure.jsonPath());
  }

  @Test
  void readsSheetManagementOperationsFromJson() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "mode": "NEW" },
              "operations": [
                { "type": "RENAME_SHEET", "sheetName": "Budget", "newSheetName": "Summary" },
                { "type": "DELETE_SHEET", "sheetName": "Scratch" },
                { "type": "MOVE_SHEET", "sheetName": "Summary", "targetIndex": 0 }
              ],
              "analysis": { "sheets": [] }
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(3, request.operations().size());
    assertInstanceOf(WorkbookOperation.RenameSheet.class, request.operations().get(0));
    assertInstanceOf(WorkbookOperation.DeleteSheet.class, request.operations().get(1));
    assertInstanceOf(WorkbookOperation.MoveSheet.class, request.operations().get(2));

    WorkbookOperation.RenameSheet renameSheet =
        (WorkbookOperation.RenameSheet) request.operations().get(0);
    WorkbookOperation.MoveSheet moveSheet =
        (WorkbookOperation.MoveSheet) request.operations().get(2);
    assertEquals("Summary", renameSheet.newSheetName());
    assertEquals(0, moveSheet.targetIndex());
  }

  @Test
  void readsStructuralLayoutOperationsFromJson() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "mode": "NEW" },
              "operations": [
                { "type": "MERGE_CELLS", "sheetName": "Budget", "range": "A1:B2" },
                { "type": "UNMERGE_CELLS", "sheetName": "Budget", "range": "A1:B2" },
                {
                  "type": "SET_COLUMN_WIDTH",
                  "sheetName": "Budget",
                  "firstColumnIndex": 0,
                  "lastColumnIndex": 2,
                  "widthCharacters": 16.0
                },
                {
                  "type": "SET_ROW_HEIGHT",
                  "sheetName": "Budget",
                  "firstRowIndex": 0,
                  "lastRowIndex": 3,
                  "heightPoints": 28.5
                },
                {
                  "type": "FREEZE_PANES",
                  "sheetName": "Budget",
                  "splitColumn": 1,
                  "splitRow": 2,
                  "leftmostColumn": 1,
                  "topRow": 2
                }
              ],
              "analysis": { "sheets": [] }
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(5, request.operations().size());
    assertInstanceOf(WorkbookOperation.MergeCells.class, request.operations().get(0));
    assertInstanceOf(WorkbookOperation.UnmergeCells.class, request.operations().get(1));
    assertInstanceOf(WorkbookOperation.SetColumnWidth.class, request.operations().get(2));
    assertInstanceOf(WorkbookOperation.SetRowHeight.class, request.operations().get(3));
    assertInstanceOf(WorkbookOperation.FreezePanes.class, request.operations().get(4));

    WorkbookOperation.SetColumnWidth setColumnWidth =
        (WorkbookOperation.SetColumnWidth) request.operations().get(2);
    WorkbookOperation.FreezePanes freezePanes =
        (WorkbookOperation.FreezePanes) request.operations().get(4);
    assertEquals(16.0, setColumnWidth.widthCharacters());
    assertEquals(2, freezePanes.topRow());
  }

  @Test
  void readsFormattingDepthStylePatchesFromJson() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "mode": "NEW" },
              "operations": [
                {
                  "type": "APPLY_STYLE",
                  "sheetName": "Budget",
                  "range": "A1",
                  "style": {
                    "bold": true,
                    "wrapText": true,
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "TOP",
                    "fontName": "Aptos",
                    "fontHeight": { "type": "POINTS", "points": 11.5 },
                    "fontColor": "#1f4e78",
                    "underline": true,
                    "strikeout": true,
                    "fillColor": "#fff2cc",
                    "border": {
                      "all": { "style": "THIN" },
                      "right": { "style": "DOUBLE" }
                    }
                  }
                }
              ],
              "analysis": { "sheets": [] }
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(1, request.operations().size());
    assertInstanceOf(WorkbookOperation.ApplyStyle.class, request.operations().getFirst());
    WorkbookOperation.ApplyStyle applyStyle =
        (WorkbookOperation.ApplyStyle) request.operations().getFirst();

    assertEquals("Aptos", applyStyle.style().fontName());
    assertEquals(
        new BigDecimal("11.5"), applyStyle.style().fontHeight().toExcelFontHeight().points());
    assertEquals("#1F4E78", applyStyle.style().fontColor());
    assertTrue(applyStyle.style().underline());
    assertTrue(applyStyle.style().strikeout());
    assertEquals("#FFF2CC", applyStyle.style().fillColor());
    assertEquals(ExcelBorderStyle.THIN, applyStyle.style().border().all().style());
    assertEquals(ExcelBorderStyle.DOUBLE, applyStyle.style().border().right().style());
  }

  @Test
  void readsExcelAuthoringOperationsAndNamedRangeAnalysisFromJson() throws IOException {
    GridGrindRequest request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "mode": "NEW" },
              "operations": [
                {
                  "type": "SET_HYPERLINK",
                  "sheetName": "Budget",
                  "address": "A1",
                  "target": { "type": "URL", "target": "https://example.com/report" }
                },
                {
                  "type": "SET_COMMENT",
                  "sheetName": "Budget",
                  "address": "A1",
                  "comment": { "text": "Review", "author": "GridGrind", "visible": true }
                },
                {
                  "type": "SET_NAMED_RANGE",
                  "name": "BudgetTotal",
                  "scope": { "type": "WORKBOOK" },
                  "target": { "sheetName": "Budget", "range": "B4" }
                }
              ],
              "analysis": {
                "sheets": [],
                "namedRanges": {
                  "mode": "SELECTED",
                  "selectors": [
                    { "type": "WORKBOOK_SCOPE", "name": "BudgetTotal" }
                  ]
                }
              }
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(3, request.operations().size());
    assertInstanceOf(WorkbookOperation.SetHyperlink.class, request.operations().get(0));
    assertInstanceOf(WorkbookOperation.SetComment.class, request.operations().get(1));
    assertInstanceOf(WorkbookOperation.SetNamedRange.class, request.operations().get(2));
    assertInstanceOf(
        GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.Selected.class,
        request.analysis().namedRanges());

    WorkbookOperation.SetHyperlink setHyperlink =
        (WorkbookOperation.SetHyperlink) request.operations().get(0);
    WorkbookOperation.SetComment setComment =
        (WorkbookOperation.SetComment) request.operations().get(1);
    WorkbookOperation.SetNamedRange setNamedRange =
        (WorkbookOperation.SetNamedRange) request.operations().get(2);
    assertEquals(
        "https://example.com/report", ((HyperlinkTarget.Url) setHyperlink.target()).target());
    assertTrue(setComment.comment().visible());
    assertEquals("BudgetTotal", setNamedRange.name());
  }

  @Test
  void wrapsRecordConstructionFailureAsInvalidPayloadError() {
    // Parsing a response where CellStyleReport.numberFormat is null triggers a NullPointerException
    // inside the record compact constructor. Jackson catches and wraps it as a JacksonException.
    InvalidJsonException failure =
        assertThrows(
            InvalidJsonException.class,
            () ->
                GridGrindJson.readResponse(
                    """
                    {"status":"SUCCESS","protocolVersion":"V1","savedWorkbookPath":null,\
                    "workbook":{"sheetCount":0,"sheetNames":[],"forceFormulaRecalculationOnOpen":false},\
                    "sheets":[{"sheetName":"X","physicalRowCount":0,"lastRowIndex":0,"lastColumnIndex":0,\
                    "requestedCells":[{"effectiveType":"BLANK","address":"A1","declaredType":"BLANK",\
                    "displayValue":"","style":{"numberFormat":null,"bold":false,"italic":false,\
                    "wrapText":false,"horizontalAlignment":"GENERAL","verticalAlignment":"BOTTOM",\
                    "fontName":"Calibri","fontHeight":{"twips":220,"points":11},"fontColor":null,"underline":false,\
                    "strikeout":false,"fillColor":null,"topBorderStyle":"NONE",\
                    "rightBorderStyle":"NONE","bottomBorderStyle":"NONE","leftBorderStyle":"NONE"}}],\
                    "previewRows":[]}]}"""
                        .getBytes(StandardCharsets.UTF_8)));

    assertNotNull(failure.getMessage());
    assertFalse(failure.getMessage().contains("[Source:"));
  }

  @Test
  void wrapsDateTimeParseFailuresAsInvalidRequestErrors() {
    // Jackson's LocalDate deserializer throws DateTimeParseException (a DateTimeException)
    // when the input string is not a valid ISO-8601 date. validationCause must detect it.
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
                      "analysis": { "sheets": [] }
                    }
                    """
                        .getBytes(java.nio.charset.StandardCharsets.UTF_8)));

    assertNotNull(failure.getMessage());
  }

  @Test
  void stripsJacksonSourceLocationFromInvalidEnumValueMessages() throws IOException {
    // An unrecognised enum value produces an InvalidFormatException whose raw message
    // contains " at [Source:" — verify the cleaned message does not expose that suffix.
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
                      "analysis": { "sheets": [] }
                    }
                    """
                        .getBytes(java.nio.charset.StandardCharsets.UTF_8)));

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
