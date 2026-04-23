package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.xml.sax.helpers.AttributesImpl;

/** Tests for low-memory workbook summary reads backed by POI's XSSF event model. */
class ExcelEventWorkbookReaderTest {
  @Test
  void matchesFullWorkbookAndSheetSummariesForSupportedReads() throws IOException {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-event-read-", ".xlsx");

    WorkbookReadResult.WorkbookSummary expectedWorkbookSummary;
    WorkbookReadResult.SheetSummary expectedSheetSummary;
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.getOrCreateSheet("Archive");
      workbook.setSheetVisibility("Archive", ExcelSheetVisibility.VERY_HIDDEN);
      workbook.setActiveSheet("Ops");
      workbook.setSelectedSheets(List.of("Ops"));
      workbook.formulas().markRecalculateOnOpen();
      workbook.sheet("Ops").setCell("A1", ExcelCellValue.text("Header"));
      workbook.sheet("Ops").setCell("D3", ExcelCellValue.number(12.5d));
      workbook.sheet("Ops").setColumnWidth(5, 5, 18.0d);
      workbook.setSheetProtection(
          "Ops",
          new ExcelSheetProtectionSettings(
              true, false, false, false, false, false, false, false, false, true, false, false,
              true, true, false),
          "secret");
      expectedWorkbookSummary = workbook.workbookSummary();
      expectedSheetSummary = workbook.sheetSummary("Ops");
      workbook.save(workbookPath);
    }

    List<WorkbookReadResult.Introspection> reads =
        new ExcelEventWorkbookReader()
            .apply(
                workbookPath,
                List.of(
                    new WorkbookReadCommand.GetWorkbookSummary("workbook"),
                    new WorkbookReadCommand.GetSheetSummary("sheet", "Ops")));

    assertEquals(
        new WorkbookReadResult.WorkbookSummaryResult("workbook", expectedWorkbookSummary),
        reads.get(0));
    assertEquals(
        new WorkbookReadResult.SheetSummaryResult("sheet", expectedSheetSummary), reads.get(1));
  }

  @Test
  void rejectsEveryUnsupportedReadCommandVariant() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-read-unsupported-", ".xlsx");
    ExcelEventWorkbookReader reader = new ExcelEventWorkbookReader();
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.save(workbookPath);
    }

    List<WorkbookReadCommand.Introspection> singleCommand = new ArrayList<>(1);
    for (WorkbookReadCommand.Introspection command :
        WorkbookSampleFixtures.introspectionCommands()) {
      if (command instanceof WorkbookReadCommand.GetWorkbookSummary
          || command instanceof WorkbookReadCommand.GetSheetSummary) {
        continue;
      }

      singleCommand.clear();
      singleCommand.add(command);
      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> reader.apply(workbookPath, singleCommand),
              () -> "expected EVENT_READ rejection for " + command.getClass().getSimpleName());

      assertTrue(failure.getMessage().contains("executionMode.readMode=EVENT_READ"));
      assertTrue(failure.getMessage().contains(expectedReadType(command)));
    }
  }

  @Test
  void reportsEmptyWorkbookSummaryWhenWorkbookXmlContainsNoSheets() throws IOException {
    Path sourceWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-empty-source-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Original");
      try (OutputStream outputStream = Files.newOutputStream(sourceWorkbook)) {
        workbook.write(outputStream);
      }
    }

    Path emptyWorkbook =
        rewriteEntry(
            sourceWorkbook,
            "xl/workbook.xml",
            xml ->
                xml.replaceAll("(?s)<bookViews>.*?</bookViews>", "")
                    .replaceAll("(?s)<sheets>.*?</sheets>", ""));

    List<WorkbookReadResult.Introspection> reads =
        new ExcelEventWorkbookReader()
            .apply(emptyWorkbook, List.of(new WorkbookReadCommand.GetWorkbookSummary("empty")));

    assertEquals(
        List.of(
            new WorkbookReadResult.WorkbookSummaryResult(
                "empty", new WorkbookReadResult.WorkbookSummary.Empty(0, List.of(), 0, false))),
        reads);
  }

  @Test
  void fallsBackToFirstSheetWhenNoTabsAreSelectedAndActiveTabIsInvalid() throws IOException {
    Path sourceWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-fallback-source-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Visible");
      workbook.createSheet("Hidden");
      workbook.setSheetHidden(1, true);
      workbook.createSheet("VeryHidden");
      workbook.setSheetVisibility(2, org.apache.poi.ss.usermodel.SheetVisibility.VERY_HIDDEN);
      workbook.createName().setNameName("BudgetTotal");
      workbook.getName("BudgetTotal").setRefersToFormula("Visible!$A$1");
      workbook.setActiveSheet(2);
      try (OutputStream outputStream = Files.newOutputStream(sourceWorkbook)) {
        workbook.write(outputStream);
      }
    }

    Path mutatedWorkbook =
        rewriteEntries(
            sourceWorkbook,
            Map.of(
                "xl/workbook.xml",
                ExcelEventWorkbookReaderTest::mutateFallbackWorkbookXml,
                "xl/worksheets/sheet1.xml",
                ExcelEventWorkbookReaderTest::clearTabSelectedAttributes,
                "xl/worksheets/sheet2.xml",
                ExcelEventWorkbookReaderTest::clearTabSelectedAttributes,
                "xl/worksheets/sheet3.xml",
                ExcelEventWorkbookReaderTest::clearTabSelectedAttributes));

    List<WorkbookReadResult.Introspection> reads =
        new ExcelEventWorkbookReader()
            .apply(
                mutatedWorkbook,
                List.of(
                    new WorkbookReadCommand.GetWorkbookSummary("workbook"),
                    new WorkbookReadCommand.GetSheetSummary("hidden", "Hidden"),
                    new WorkbookReadCommand.GetSheetSummary("veryHidden", "VeryHidden")));

    assertEquals(
        new WorkbookReadResult.WorkbookSummaryResult(
            "workbook",
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                3,
                List.of("Visible", "Hidden", "VeryHidden"),
                "Visible",
                List.of("Visible"),
                1,
                false)),
        reads.get(0));
    assertEquals(
        new WorkbookReadResult.SheetSummaryResult(
            "hidden",
            new WorkbookReadResult.SheetSummary(
                "Hidden",
                ExcelSheetVisibility.HIDDEN,
                new WorkbookReadResult.SheetProtection.Unprotected(),
                0,
                -1,
                -1)),
        reads.get(1));
    assertEquals(
        new WorkbookReadResult.SheetSummaryResult(
            "veryHidden",
            new WorkbookReadResult.SheetSummary(
                "VeryHidden",
                ExcelSheetVisibility.VERY_HIDDEN,
                new WorkbookReadResult.SheetProtection.Unprotected(),
                0,
                -1,
                -1)),
        reads.get(2));
  }

  @Test
  void fallsBackToFirstSheetWhenWorkbookXmlUsesNegativeActiveTabAndBareCalcPr() throws IOException {
    Path sourceWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-negative-active-tab-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Visible");
      workbook.createSheet("Other");
      workbook.setActiveSheet(1);
      try (OutputStream outputStream = Files.newOutputStream(sourceWorkbook)) {
        workbook.write(outputStream);
      }
    }

    Path mutatedWorkbook =
        rewriteEntry(
            sourceWorkbook,
            "xl/workbook.xml",
            xml -> {
              String withNegativeActiveTab =
                  xml.contains("activeTab=")
                      ? xml.replaceAll("activeTab=\"[^\"]*\"", "activeTab=\"-1\"")
                      : xml.replace("<workbookView ", "<workbookView activeTab=\"-1\" ");
              return withNegativeActiveTab.contains("<calcPr")
                  ? withNegativeActiveTab.replaceAll("<calcPr[^>]*/>", "<calcPr/>")
                  : withNegativeActiveTab.replace("</workbook>", "<calcPr/></workbook>");
            });

    WorkbookReadResult.WorkbookSummary.WithSheets summary =
        assertInstanceOf(
            WorkbookReadResult.WorkbookSummary.WithSheets.class,
            ((WorkbookReadResult.WorkbookSummaryResult)
                    new ExcelEventWorkbookReader()
                        .apply(
                            mutatedWorkbook,
                            List.of(new WorkbookReadCommand.GetWorkbookSummary("workbook")))
                        .getFirst())
                .workbook());

    assertEquals("Visible", summary.activeSheetName());
    assertFalse(summary.forceFormulaRecalculationOnOpen());
  }

  @Test
  void reportsSparseRowAndColumnFactsFromCustomizedSheetXml() throws IOException {
    Path sourceWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-sparse-source-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Sparse");
      try (OutputStream outputStream = Files.newOutputStream(sourceWorkbook)) {
        workbook.write(outputStream);
      }
    }

    Path sparseWorkbook =
        rewriteEntry(
            sourceWorkbook,
            "xl/worksheets/sheet1.xml",
            _ ->
                """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetViews>
                    <sheetView workbookViewId="0"/>
                  </sheetViews>
                  <cols>
                    <col min="1" max=""/>
                  </cols>
                  <sheetData>
                    <row>
                      <c><v>1</v></c>
                    </row>
                    <row r="4">
                      <c r="C4"><v>2</v></c>
                    </row>
                  </sheetData>
                  <sheetProtection/>
                </worksheet>
                """);

    List<WorkbookReadResult.Introspection> reads =
        new ExcelEventWorkbookReader()
            .apply(
                sparseWorkbook,
                List.of(new WorkbookReadCommand.GetSheetSummary("sheet", "Sparse")));

    assertEquals(
        List.of(
            new WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new WorkbookReadResult.SheetSummary(
                    "Sparse",
                    ExcelSheetVisibility.VISIBLE,
                    new WorkbookReadResult.SheetProtection.Unprotected(),
                    2,
                    3,
                    2))),
        reads);
  }

  @Test
  void throwsWhenSheetSummaryTargetsMissingSheet() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-missing-sheet-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.save(workbookPath);
    }

    SheetNotFoundException failure =
        assertThrows(
            SheetNotFoundException.class,
            () ->
                new ExcelEventWorkbookReader()
                    .apply(
                        workbookPath,
                        List.of(new WorkbookReadCommand.GetSheetSummary("missing", "Absent"))));

    assertEquals("Sheet does not exist: Absent", failure.getMessage());
  }

  @Test
  void rejectsNonXlsxAndWrapsMalformedWorkbookAndSheetXmlAsIoExceptions() throws IOException {
    Path invalidWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-invalid-", ".xlsx");
    Files.writeString(invalidWorkbook, "not a workbook", StandardCharsets.UTF_8);

    IllegalArgumentException workbookFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ExcelEventWorkbookReader()
                    .apply(
                        invalidWorkbook,
                        List.of(new WorkbookReadCommand.GetWorkbookSummary("workbook"))));
    assertEquals("Only .xlsx workbooks are supported", workbookFailure.getMessage());

    Path sourceWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-bad-workbook-source-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.save(sourceWorkbook);
    }
    Path malformedWorkbook = rewriteEntry(sourceWorkbook, "xl/workbook.xml", _ -> "<workbook");

    IOException metadataFailure =
        assertThrows(
            IOException.class,
            () ->
                new ExcelEventWorkbookReader()
                    .apply(
                        malformedWorkbook,
                        List.of(new WorkbookReadCommand.GetWorkbookSummary("workbook"))));
    assertTrue(
        metadataFailure
            .getMessage()
            .contains("Failed to read workbook through the XSSF event model"));

    Path malformedSheetSource =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-bad-sheet-source-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.save(malformedSheetSource);
    }
    Path malformedSheetWorkbook =
        rewriteEntry(malformedSheetSource, "xl/worksheets/sheet1.xml", _ -> "<worksheet");

    IOException sheetFailure =
        assertThrows(
            IOException.class,
            () ->
                new ExcelEventWorkbookReader()
                    .apply(
                        malformedSheetWorkbook,
                        List.of(new WorkbookReadCommand.GetSheetSummary("sheet", "Ops"))));

    assertTrue(sheetFailure.getMessage().contains("Failed to parse sheet Ops"));
  }

  @Test
  void wrapsNonSaxSheetParseFailuresWithSheetContext() throws IOException {
    Path sourceWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-invalid-row-source-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.save(sourceWorkbook);
    }
    Path malformedSheetWorkbook =
        rewriteEntry(
            sourceWorkbook,
            "xl/worksheets/sheet1.xml",
            _ ->
                """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="nope">
                      <c r="A1"><v>1</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """);

    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                new ExcelEventWorkbookReader()
                    .apply(
                        malformedSheetWorkbook,
                        List.of(new WorkbookReadCommand.GetSheetSummary("sheet", "Ops"))));

    assertTrue(failure.getMessage().contains("Failed to parse sheet Ops"));
  }

  @Test
  void helperSeamsCoverSupportedSummaryCommandTypesAndEmptyMetadataGuard() {
    assertEquals(
        "GET_WORKBOOK_SUMMARY",
        EventReadCommandTypes.commandType(new WorkbookReadCommand.GetWorkbookSummary("workbook")));
    assertEquals(
        "GET_SHEET_SUMMARY",
        EventReadCommandTypes.commandType(new WorkbookReadCommand.GetSheetSummary("sheet", "Ops")));
    assertEquals(
        "GET_PACKAGE_SECURITY",
        EventReadCommandTypes.commandType(new WorkbookReadCommand.GetPackageSecurity("security")));
    assertEquals(7, invokeActiveSheetIndexWithExplicitActiveTab());
    assertEquals(0, invokeActiveSheetIndexWithoutActiveTab());
    assertEquals(0, invokeActiveSheetIndexWithEmptyBookViews());

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                EventWorkbookMetadata.activeSheetName(
                    new EventWorkbookMetadata(List.of(), Map.of(), 0, 0, false)));
    assertEquals("workbook metadata must contain at least one sheet", failure.getMessage());
  }

  @Test
  void sheetSummaryHandlerFallsBackToQNameWhenLocalNameIsEmpty() {
    EventSheetSummaryHandler handler = new EventSheetSummaryHandler();

    AttributesImpl attributes = new AttributesImpl();
    attributes.addAttribute("", "r", "r", "CDATA", "2");
    handler.startElement("", "", "row", attributes);

    assertEquals(1, handler.summary().lastRowIndex());
  }

  @Test
  void sheetSummaryHandlerAcceptsNullLocalNamesAndTrueAttributesCaseInsensitively() {
    EventSheetSummaryHandler handler = new EventSheetSummaryHandler();

    AttributesImpl selectedAttributes = new AttributesImpl();
    selectedAttributes.addAttribute("", "tabSelected", "tabSelected", "CDATA", "TRUE");
    handler.startElement("", null, "sheetView", selectedAttributes);

    handler.startElement("", null, "sheetView", new AttributesImpl());

    AttributesImpl protectionAttributes = new AttributesImpl();
    protectionAttributes.addAttribute("", "sheet", "sheet", "CDATA", "1");
    handler.startElement("", null, "sheetProtection", protectionAttributes);

    EventSheetSummary summary = handler.summary();
    assertTrue(summary.selected());
    assertInstanceOf(WorkbookReadResult.SheetProtection.Protected.class, summary.protection());
  }

  private static String expectedReadType(WorkbookReadCommand.Introspection command) {
    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadCommand.GetWorkbookProtection _ -> "GET_WORKBOOK_PROTECTION";
      case WorkbookReadCommand.GetCustomXmlMappings _ -> "GET_CUSTOM_XML_MAPPINGS";
      case WorkbookReadCommand.ExportCustomXmlMapping _ -> "EXPORT_CUSTOM_XML_MAPPING";
      case WorkbookReadCommand.GetNamedRanges _ -> "GET_NAMED_RANGES";
      case WorkbookReadCommand.GetSheetSummary _ -> "GET_SHEET_SUMMARY";
      case WorkbookReadCommand.GetCells _ -> "GET_CELLS";
      case WorkbookReadCommand.GetWindow _ -> "GET_WINDOW";
      case WorkbookReadCommand.GetMergedRegions _ -> "GET_MERGED_REGIONS";
      case WorkbookReadCommand.GetHyperlinks _ -> "GET_HYPERLINKS";
      case WorkbookReadCommand.GetComments _ -> "GET_COMMENTS";
      case WorkbookReadCommand.GetDrawingObjects _ -> "GET_DRAWING_OBJECTS";
      case WorkbookReadCommand.GetCharts _ -> "GET_CHARTS";
      case WorkbookReadCommand.GetPivotTables _ -> "GET_PIVOT_TABLES";
      case WorkbookReadCommand.GetDrawingObjectPayload _ -> "GET_DRAWING_OBJECT_PAYLOAD";
      case WorkbookReadCommand.GetSheetLayout _ -> "GET_SHEET_LAYOUT";
      case WorkbookReadCommand.GetPrintLayout _ -> "GET_PRINT_LAYOUT";
      case WorkbookReadCommand.GetDataValidations _ -> "GET_DATA_VALIDATIONS";
      case WorkbookReadCommand.GetConditionalFormatting _ -> "GET_CONDITIONAL_FORMATTING";
      case WorkbookReadCommand.GetAutofilters _ -> "GET_AUTOFILTERS";
      case WorkbookReadCommand.GetTables _ -> "GET_TABLES";
      case WorkbookReadCommand.GetArrayFormulas _ -> "GET_ARRAY_FORMULAS";
      case WorkbookReadCommand.GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadCommand.GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadCommand.GetPackageSecurity _ -> "GET_PACKAGE_SECURITY";
      case WorkbookReadCommand.GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
    };
  }

  private static Path rewriteEntry(
      Path sourceWorkbook, String entryName, java.util.function.UnaryOperator<String> transformer)
      throws IOException {
    return OoxmlPartMutator.rewriteEntry(sourceWorkbook, entryName, transformer);
  }

  private static Path rewriteEntries(
      Path sourceWorkbook, Map<String, java.util.function.UnaryOperator<String>> transformers)
      throws IOException {
    return OoxmlPartMutator.rewriteEntries(sourceWorkbook, transformers);
  }

  @Test
  void readsEdgeCasesInWorkbookXmlAndSheetXml() throws IOException {
    Path sourceWorkbook =
        ExcelTempFiles.createManagedTempFile("gridgrind-event-edge-source-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Visible");
      workbook.createSheet("Other");
      workbook.createName().setNameName("CounterFn");
      workbook.getName("CounterFn").setRefersToFormula("Visible!$A$1");
      workbook.createName().setNameName("CounterRange");
      workbook.getName("CounterRange").setRefersToFormula("Visible!$A$2");
      try (OutputStream outputStream = Files.newOutputStream(sourceWorkbook)) {
        workbook.write(outputStream);
      }
    }

    // Mutate workbook.xml to cover edge-case branches in workbookMetadata and namedRangeCount,
    // then mutate the sheet XML to cover edge-case branches in SheetSummaryHandler.
    Path edgeCaseWorkbook =
        rewriteEntries(
            sourceWorkbook,
            Map.of(
                "xl/workbook.xml",
                xml -> {
                  // state="visible" explicitly set — covers visibility() VISIBLE true branch.
                  String withExplicitVisible =
                      xml.replaceFirst("(<sheet[^/]+ name=\"Visible\")", "$1 state=\"visible\"");
                  // calcPr with fullCalcOnLoad="0" — covers isSetFullCalcOnLoad()=true,
                  // getFullCalcOnLoad()=false branch.
                  String withCalcPr =
                      withExplicitVisible.contains("calcPr")
                          ? withExplicitVisible.replaceAll(
                              "<calcPr[^/]*/?>", "<calcPr fullCalcOnLoad=\"0\"/>")
                          : withExplicitVisible.replace(
                              "</workbook>", "<calcPr fullCalcOnLoad=\"0\"/></workbook>");
                  // Named ranges with function="0" and hidden="0" — covers isSetFunction=true,
                  // getFunction=false and isSetHidden=true, getHidden=false branches.
                  return withCalcPr.replace(
                      "name=\"CounterFn\"", "name=\"CounterFn\" function=\"0\" hidden=\"0\"");
                },
                "xl/worksheets/sheet1.xml",
                _ ->
                    """
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <cols>
                        <col min="1"/>
                      </cols>
                      <sheetData>
                        <row r="1">
                          <c r=""><v>1</v></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """));

    List<WorkbookReadResult.Introspection> reads =
        new ExcelEventWorkbookReader()
            .apply(
                edgeCaseWorkbook,
                List.of(
                    new WorkbookReadCommand.GetWorkbookSummary("workbook"),
                    new WorkbookReadCommand.GetSheetSummary("sheet", "Visible")));

    WorkbookReadResult.WorkbookSummary.WithSheets summary =
        assertInstanceOf(
            WorkbookReadResult.WorkbookSummary.WithSheets.class,
            ((WorkbookReadResult.WorkbookSummaryResult) reads.get(0)).workbook());
    assertFalse(summary.forceFormulaRecalculationOnOpen());
    WorkbookReadResult.SheetSummary sheetSummary =
        ((WorkbookReadResult.SheetSummaryResult) reads.get(1)).sheet();
    assertEquals(ExcelSheetVisibility.VISIBLE, sheetSummary.visibility());
    assertEquals(1, sheetSummary.physicalRowCount());
    assertEquals(-1, sheetSummary.lastColumnIndex());
  }

  private static int invokeActiveSheetIndexWithEmptyBookViews() {
    CTWorkbook workbook = CTWorkbook.Factory.newInstance();
    workbook.addNewBookViews(); // bookViews present but contains no workbookView entries
    return EventWorkbookMetadata.activeSheetIndex(workbook);
  }

  private static String clearTabSelectedAttributes(String xml) {
    return xml.replaceAll("\\s+tabSelected=\"[^\"]*\"", "");
  }

  private static String mutateFallbackWorkbookXml(String xml) {
    String withInvalidActiveTab =
        xml.contains("activeTab=")
            ? xml.replaceAll("activeTab=\"[^\"]*\"", "activeTab=\"99\"")
            : xml.replace("<workbookView ", "<workbookView activeTab=\"99\" ");
    return withInvalidActiveTab.replace(
        "</definedNames>",
        "<definedName name=\"HiddenFn\" function=\"1\" hidden=\"1\">Visible!$A$1</definedName>"
            + "</definedNames>");
  }

  private static int invokeActiveSheetIndexWithExplicitActiveTab() {
    CTWorkbook workbook = CTWorkbook.Factory.newInstance();
    workbook.addNewBookViews().addNewWorkbookView().setActiveTab(7);
    return EventWorkbookMetadata.activeSheetIndex(workbook);
  }

  private static int invokeActiveSheetIndexWithoutActiveTab() {
    CTWorkbook workbook = CTWorkbook.Factory.newInstance();
    workbook.addNewBookViews().addNewWorkbookView();
    return EventWorkbookMetadata.activeSheetIndex(workbook);
  }
}
