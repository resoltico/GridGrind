package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.xml.sax.helpers.AttributesImpl;

/** Tests for low-memory workbook summary reads backed by POI's XSSF event model. */
@SuppressWarnings("PMD.AvoidAccessibilityAlteration")
class ExcelEventWorkbookReaderTest {
  @Test
  void matchesFullWorkbookAndSheetSummariesForSupportedReads() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-event-read-", ".xlsx");

    WorkbookReadResult.WorkbookSummary expectedWorkbookSummary;
    WorkbookReadResult.SheetSummary expectedSheetSummary;
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.getOrCreateSheet("Archive");
      workbook.setSheetVisibility("Archive", ExcelSheetVisibility.VERY_HIDDEN);
      workbook.setActiveSheet("Ops");
      workbook.setSelectedSheets(List.of("Ops"));
      workbook.forceFormulaRecalculationOnOpen();
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
    Path workbookPath = Files.createTempFile("gridgrind-event-read-unsupported-", ".xlsx");
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
    Path sourceWorkbook = Files.createTempFile("gridgrind-event-empty-source-", ".xlsx");
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
    Path sourceWorkbook = Files.createTempFile("gridgrind-event-fallback-source-", ".xlsx");
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
  void reportsSparseRowAndColumnFactsFromCustomizedSheetXml() throws IOException {
    Path sourceWorkbook = Files.createTempFile("gridgrind-event-sparse-source-", ".xlsx");
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
    Path workbookPath = Files.createTempFile("gridgrind-event-missing-sheet-", ".xlsx");
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
    Path invalidWorkbook = Files.createTempFile("gridgrind-event-invalid-", ".xlsx");
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

    Path sourceWorkbook = Files.createTempFile("gridgrind-event-bad-workbook-source-", ".xlsx");
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

    Path malformedSheetSource = Files.createTempFile("gridgrind-event-bad-sheet-source-", ".xlsx");
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
    Path sourceWorkbook = Files.createTempFile("gridgrind-event-invalid-row-source-", ".xlsx");
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
  void reflectiveHelpersCoverSupportedSummaryCommandTypesAndEmptyMetadataGuard() throws Exception {
    assertEquals(
        "GET_WORKBOOK_SUMMARY",
        invokeCommandType(new WorkbookReadCommand.GetWorkbookSummary("workbook")));
    assertEquals(
        "GET_SHEET_SUMMARY",
        invokeCommandType(new WorkbookReadCommand.GetSheetSummary("sheet", "Ops")));
    assertEquals(
        "GET_PACKAGE_SECURITY",
        invokeCommandType(new WorkbookReadCommand.GetPackageSecurity("security")));
    assertEquals(7, invokeActiveSheetIndexWithExplicitActiveTab());
    assertEquals(0, invokeActiveSheetIndexWithoutActiveTab());

    Constructor<?> constructor =
        Class.forName("dev.erst.gridgrind.excel.ExcelEventWorkbookReader$EventWorkbookMetadata")
            .getDeclaredConstructor(List.class, Map.class, int.class, int.class, boolean.class);
    constructor.setAccessible(true);
    Object metadata = constructor.newInstance(List.of(), Map.of(), 0, 0, false);

    Method activeSheetName =
        ExcelEventWorkbookReader.class.getDeclaredMethod("activeSheetName", metadata.getClass());
    activeSheetName.setAccessible(true);

    InvocationTargetException failure =
        assertThrows(InvocationTargetException.class, () -> activeSheetName.invoke(null, metadata));
    assertInstanceOf(IllegalStateException.class, failure.getCause());
    assertEquals(
        "workbook metadata must contain at least one sheet", failure.getCause().getMessage());
  }

  @Test
  void sheetSummaryHandlerFallsBackToQNameWhenLocalNameIsEmpty() throws Exception {
    Constructor<?> constructor =
        Class.forName("dev.erst.gridgrind.excel.ExcelEventWorkbookReader$SheetSummaryHandler")
            .getDeclaredConstructor();
    constructor.setAccessible(true);
    Object handler = constructor.newInstance();

    Method startElement =
        handler
            .getClass()
            .getDeclaredMethod(
                "startElement",
                String.class,
                String.class,
                String.class,
                org.xml.sax.Attributes.class);
    startElement.setAccessible(true);

    AttributesImpl attributes = new AttributesImpl();
    attributes.addAttribute("", "r", "r", "CDATA", "2");
    startElement.invoke(handler, "", "", "row", attributes);

    Method summary = handler.getClass().getDeclaredMethod("summary");
    summary.setAccessible(true);
    Object eventSheetSummary = summary.invoke(handler);
    Method lastRowIndex = eventSheetSummary.getClass().getDeclaredMethod("lastRowIndex");
    lastRowIndex.setAccessible(true);

    assertEquals(1, lastRowIndex.invoke(eventSheetSummary));
  }

  private static String expectedReadType(WorkbookReadCommand.Introspection command) {
    return switch (command) {
      case WorkbookReadCommand.GetWorkbookSummary _ -> "GET_WORKBOOK_SUMMARY";
      case WorkbookReadCommand.GetWorkbookProtection _ -> "GET_WORKBOOK_PROTECTION";
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
      case WorkbookReadCommand.GetFormulaSurface _ -> "GET_FORMULA_SURFACE";
      case WorkbookReadCommand.GetSheetSchema _ -> "GET_SHEET_SCHEMA";
      case WorkbookReadCommand.GetNamedRangeSurface _ -> "GET_NAMED_RANGE_SURFACE";
      case WorkbookReadCommand.GetPackageSecurity _ -> "GET_PACKAGE_SECURITY";
    };
  }

  private static Path rewriteEntry(
      Path sourceWorkbook, String entryName, TextTransformer transformer) throws IOException {
    return rewriteEntries(sourceWorkbook, Map.of(entryName, transformer));
  }

  private static Path rewriteEntries(Path sourceWorkbook, Map<String, TextTransformer> transformers)
      throws IOException {
    List<ZipEntryBytes> entries = new ArrayList<>();
    try (ZipFile zipFile = new ZipFile(sourceWorkbook.toFile())) {
      Enumeration<? extends ZipEntry> enumeration = zipFile.entries();
      while (enumeration.hasMoreElements()) {
        ZipEntry entry = enumeration.nextElement();
        try (InputStream inputStream = zipFile.getInputStream(entry)) {
          byte[] bytes = inputStream.readAllBytes();
          TextTransformer transformer = transformers.get(entry.getName());
          entries.add(
              new ZipEntryBytes(entry.getName(), transformedEntryBytes(bytes, transformer)));
        }
      }
    }

    Path mutatedWorkbook = Files.createTempFile("gridgrind-event-mutated-", ".xlsx");
    try (ZipOutputStream outputStream =
        new ZipOutputStream(Files.newOutputStream(mutatedWorkbook))) {
      for (ZipEntryBytes entry : entries) {
        writeZipEntry(outputStream, entry.name(), entry.bytes());
      }
    }
    return mutatedWorkbook;
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

  private static String invokeCommandType(WorkbookReadCommand.Introspection command)
      throws ReflectiveOperationException {
    Method commandType =
        ExcelEventWorkbookReader.class.getDeclaredMethod(
            "commandType", WorkbookReadCommand.Introspection.class);
    commandType.setAccessible(true);
    return (String) commandType.invoke(null, command);
  }

  private static int invokeActiveSheetIndexWithExplicitActiveTab()
      throws ReflectiveOperationException {
    CTWorkbook workbook = CTWorkbook.Factory.newInstance();
    workbook.addNewBookViews().addNewWorkbookView().setActiveTab(7);

    Method activeSheetIndex =
        ExcelEventWorkbookReader.class.getDeclaredMethod("activeSheetIndex", CTWorkbook.class);
    activeSheetIndex.setAccessible(true);
    return (int) activeSheetIndex.invoke(null, workbook);
  }

  private static int invokeActiveSheetIndexWithoutActiveTab() throws ReflectiveOperationException {
    CTWorkbook workbook = CTWorkbook.Factory.newInstance();
    workbook.addNewBookViews().addNewWorkbookView();

    Method activeSheetIndex =
        ExcelEventWorkbookReader.class.getDeclaredMethod("activeSheetIndex", CTWorkbook.class);
    activeSheetIndex.setAccessible(true);
    return (int) activeSheetIndex.invoke(null, workbook);
  }

  private static byte[] transformedEntryBytes(byte[] bytes, TextTransformer transformer) {
    if (transformer == null) {
      return bytes;
    }
    return transformer
        .transform(new String(bytes, StandardCharsets.UTF_8))
        .getBytes(StandardCharsets.UTF_8);
  }

  private static void writeZipEntry(ZipOutputStream outputStream, String name, byte[] bytes)
      throws IOException {
    outputStream.putNextEntry(new ZipEntry(name));
    outputStream.write(bytes);
    outputStream.closeEntry();
  }

  /** String-to-string zip-entry transformer used by the temporary workbook mutation helpers. */
  @FunctionalInterface
  private interface TextTransformer {
    String transform(String xml);
  }

  /** Immutable zip-entry payload used to preserve original OOXML part order in test mutations. */
  private static final class ZipEntryBytes {
    private final String name;
    private final byte[] bytes;

    private ZipEntryBytes(String name, byte[] bytes) {
      this.name = name;
      this.bytes = bytes.clone();
    }

    private String name() {
      return name;
    }

    private byte[] bytes() {
      return bytes.clone();
    }
  }
}
