package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.InspectionIntrospectionResult;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.eventread.EventSheetSummary;
import dev.erst.gridgrind.engine.backend.eventread.EventSheetSummaryHandler;
import dev.erst.gridgrind.engine.backend.eventread.EventWorkbookMetadata;
import dev.erst.gridgrind.engine.backend.eventread.EventWorkbookReader;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.xml.sax.helpers.AttributesImpl;

/** Tests for low-memory workbook summary reads backed by POI's XSSF event model. */
class ExcelEventWorkbookReaderTest {
  @Test
  void matchesFullWorkbookAndSheetSummariesForSupportedReads() throws IOException {
    Path workbookPath = ManagedTempFiles.createManagedTempFile("gridgrind-event-read-", ".xlsx");

    WorkbookSummary expectedWorkbookSummary;
    SheetSummaryReport expectedSheetSummary;
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.getOrCreateSheet("Archive");
      workbook.sheets().setSheetVisibility("Archive", ExcelSheetVisibility.VERY_HIDDEN);
      workbook.sheets().setActiveSheet("Ops");
      workbook.sheets().setSelectedSheets(List.of("Ops"));
      workbook.formulas().markRecalculateOnOpen();
      workbook.sheet("Ops").cells().setCell("A1", new CellValue.Text("Header"));
      workbook.sheet("Ops").cells().setCell("D3", new CellValue.NumberValue(12.5d));
      workbook.sheet("Ops").columns().setWidth(5, 5, 18.0d);
      workbook
          .protection()
          .setSheetProtection(
              "Ops",
              new SheetProtectionSettings(
                  true, false, false, false, false, false, false, false, false, true, false, false,
                  true, true, false),
              Optional.of("secret"));
      expectedWorkbookSummary = workbook.workbookSummary();
      expectedSheetSummary = workbook.sheetSummary("Ops");
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    List<InspectionIntrospectionResult> reads =
        new EventWorkbookReader()
            .apply(
                workbookPath,
                List.of(
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("workbook"),
                        new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                        new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                            .GetWorkbookSummary()),
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("sheet"),
                        new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetSummary()),
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("sheet-again"),
                        new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                            .GetSheetSummary())));

    assertEquals(
        new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult(
            new StepId("workbook"), expectedWorkbookSummary),
        reads.get(0));
    assertEquals(
        new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
            new StepId("sheet"), expectedSheetSummary),
        reads.get(1));
    assertEquals(
        new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
            new StepId("sheet-again"), expectedSheetSummary),
        reads.get(2));
  }

  @Test
  void rejectsEveryUnsupportedReadCommandVariant() throws IOException {
    Path workbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-read-unsupported-", ".xlsx");
    EventWorkbookReader reader = new EventWorkbookReader();
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    List<dev.erst.gridgrind.core.plan.ResolvedInspectionStep> singleCommand = new ArrayList<>(1);
    for (dev.erst.gridgrind.core.plan.ResolvedInspectionStep command :
        WorkbookSampleFixtures.introspectionCommands()) {
      if (command.query()
              instanceof dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery.GetWorkbookSummary
          || command.query()
              instanceof dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetSummary) {
        continue;
      }

      singleCommand.clear();
      singleCommand.add(command);
      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> reader.apply(workbookPath, singleCommand),
              () -> "expected EVENT_READ rejection for " + command.getClass().getSimpleName());

      assertTrue(failure.getMessage().contains("execution.mode.type=EVENT_READ"));
      assertTrue(failure.getMessage().contains(expectedReadType(command)));
    }
  }

  @Test
  void reportsEmptyWorkbookSummaryWhenWorkbookXmlContainsNoSheets() throws IOException {
    Path sourceWorkbook =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-empty-source-", ".xlsx");
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

    List<InspectionIntrospectionResult> reads =
        new EventWorkbookReader()
            .apply(
                emptyWorkbook,
                List.of(
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("empty"),
                        new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                        new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                            .GetWorkbookSummary())));

    assertEquals(
        List.of(
            new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult(
                new StepId("empty"), new WorkbookSummary.Empty(0, List.of(), 0, false))),
        reads);
  }

  @Test
  void fallsBackToFirstSheetWhenNoTabsAreSelectedAndActiveTabIsInvalid() throws IOException {
    Path sourceWorkbook =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-fallback-source-", ".xlsx");
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

    List<InspectionIntrospectionResult> reads =
        new EventWorkbookReader()
            .apply(
                mutatedWorkbook,
                List.of(
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("workbook"),
                        new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                        new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                            .GetWorkbookSummary()),
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("hidden"),
                        new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Hidden")),
                        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetSheetSummary()),
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("veryHidden"),
                        new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                            new dev.erst.gridgrind.core.vocabulary.SheetName("VeryHidden")),
                        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                            .GetSheetSummary())));

    assertEquals(
        new dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult(
            new StepId("workbook"),
            new WorkbookSummary.WithSheets(
                3,
                List.of("Visible", "Hidden", "VeryHidden").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Visible"),
                List.of("Visible").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
                    .toList(),
                1,
                false)),
        reads.get(0));
    assertEquals(
        new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
            new StepId("hidden"),
            new SheetSummaryReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Hidden"),
                ExcelSheetVisibility.HIDDEN,
                new SheetProtectionReport.Unprotected(),
                0,
                -1,
                -1)),
        reads.get(1));
    assertEquals(
        new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
            new StepId("veryHidden"),
            new SheetSummaryReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("VeryHidden"),
                ExcelSheetVisibility.VERY_HIDDEN,
                new SheetProtectionReport.Unprotected(),
                0,
                -1,
                -1)),
        reads.get(2));
  }

  @Test
  void fallsBackToFirstSheetWhenWorkbookXmlUsesNegativeActiveTabAndBareCalcPr() throws IOException {
    Path sourceWorkbook =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-negative-active-tab-", ".xlsx");
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

    WorkbookSummary.WithSheets summary =
        assertInstanceOf(
            WorkbookSummary.WithSheets.class,
            ((dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult)
                    new EventWorkbookReader()
                        .apply(
                            mutatedWorkbook,
                            List.of(
                                new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                    new StepId("workbook"),
                                    new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                                    new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                                        .GetWorkbookSummary())))
                        .getFirst())
                .workbook());

    assertEquals(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Visible"), summary.activeSheetName());
    assertFalse(summary.forceFormulaRecalculationOnOpen());
  }

  @Test
  void reportsSparseRowAndColumnFactsFromCustomizedSheetXml() throws IOException {
    Path sourceWorkbook =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-sparse-source-", ".xlsx");
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

    List<InspectionIntrospectionResult> reads =
        new EventWorkbookReader()
            .apply(
                sparseWorkbook,
                List.of(
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("sheet"),
                        new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Sparse")),
                        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                            .GetSheetSummary())));

    assertEquals(
        List.of(
            new dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult(
                new StepId("sheet"),
                new SheetSummaryReport(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Sparse"),
                    ExcelSheetVisibility.VISIBLE,
                    new SheetProtectionReport.Unprotected(),
                    2,
                    3,
                    2))),
        reads);
  }

  @Test
  void throwsWhenSheetSummaryTargetsMissingSheet() throws IOException {
    Path workbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-missing-sheet-", ".xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    SheetNotFoundException failure =
        assertThrows(
            SheetNotFoundException.class,
            () ->
                new EventWorkbookReader()
                    .apply(
                        workbookPath,
                        List.of(
                            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                new StepId("missing"),
                                new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                                    new dev.erst.gridgrind.core.vocabulary.SheetName("Absent")),
                                new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                                    .GetSheetSummary()))));

    assertEquals("Sheet does not exist: Absent", failure.getMessage());
  }

  @Test
  void rejectsNonXlsxAndWrapsMalformedWorkbookAndSheetXmlAsIoExceptions() throws IOException {
    Path invalidWorkbook =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-invalid-", ".xlsx");
    Files.writeString(invalidWorkbook, "not a workbook", StandardCharsets.UTF_8);

    IllegalArgumentException workbookFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new EventWorkbookReader()
                    .apply(
                        invalidWorkbook,
                        List.of(
                            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                new StepId("workbook"),
                                new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                                new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                                    .GetWorkbookSummary()))));
    assertEquals("Only .xlsx workbooks are supported", workbookFailure.getMessage());

    Path sourceWorkbook =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-bad-workbook-source-", ".xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .persistence()
          .save(
              sourceWorkbook,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    Path malformedWorkbook = rewriteEntry(sourceWorkbook, "xl/workbook.xml", _ -> "<workbook");

    IOException metadataFailure =
        assertThrows(
            IOException.class,
            () ->
                new EventWorkbookReader()
                    .apply(
                        malformedWorkbook,
                        List.of(
                            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                new StepId("workbook"),
                                new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                                new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                                    .GetWorkbookSummary()))));
    assertTrue(
        metadataFailure
            .getMessage()
            .contains("Failed to read workbook through the XSSF event model"));

    Path malformedSheetSource =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-bad-sheet-source-", ".xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .persistence()
          .save(
              malformedSheetSource,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    Path malformedSheetWorkbook =
        rewriteEntry(malformedSheetSource, "xl/worksheets/sheet1.xml", _ -> "<worksheet");

    IOException sheetFailure =
        assertThrows(
            IOException.class,
            () ->
                new EventWorkbookReader()
                    .apply(
                        malformedSheetWorkbook,
                        List.of(
                            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                new StepId("sheet"),
                                new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                                    new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                                new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                                    .GetSheetSummary()))));

    assertTrue(sheetFailure.getMessage().contains("Failed to parse sheet Ops"));
  }

  @Test
  void wrapsNonSaxSheetParseFailuresWithSheetContext() throws IOException {
    Path sourceWorkbook =
        ManagedTempFiles.createManagedTempFile("gridgrind-event-invalid-row-source-", ".xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook
          .persistence()
          .save(
              sourceWorkbook,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
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
                new EventWorkbookReader()
                    .apply(
                        malformedSheetWorkbook,
                        List.of(
                            new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                                new StepId("sheet"),
                                new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                                    new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                                new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                                    .GetSheetSummary()))));

    assertTrue(failure.getMessage().contains("Failed to parse sheet Ops"));
  }

  @Test
  void helperSeamsCoverSupportedSummaryCommandTypesAndEmptyMetadataGuard() {
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
    assertInstanceOf(SheetProtectionReport.Protected.class, summary.protection());
  }

  private static String expectedReadType(
      dev.erst.gridgrind.core.plan.ResolvedInspectionStep command) {
    return command.query().queryType();
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
        ManagedTempFiles.createManagedTempFile("gridgrind-event-edge-source-", ".xlsx");
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

    List<InspectionIntrospectionResult> reads =
        new EventWorkbookReader()
            .apply(
                edgeCaseWorkbook,
                List.of(
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("workbook"),
                        new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                        new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                            .GetWorkbookSummary()),
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId("sheet"),
                        new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                            new dev.erst.gridgrind.core.vocabulary.SheetName("Visible")),
                        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                            .GetSheetSummary())));

    WorkbookSummary.WithSheets summary =
        assertInstanceOf(
            WorkbookSummary.WithSheets.class,
            ((dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult)
                    reads.get(0))
                .workbook());
    assertFalse(summary.forceFormulaRecalculationOnOpen());
    SheetSummaryReport sheetSummary =
        ((dev.erst.gridgrind.core.plan.SheetInspectionResult.SheetSummaryResult) reads.get(1))
            .sheet();
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
