package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for ExcelWorkbookIntrospector workbook-fact reads and named-range selection. */
class ExcelWorkbookIntrospectorTest {
  @Test
  void executesEveryIntrospectionCommandAgainstWorkbookState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populateWorkbookForFullIntrospection(workbook);

      IntrospectionReadResults results =
          readEveryIntrospectionResult(workbook, new ExcelWorkbookIntrospector());

      assertEquals(List.of("Budget", "Ops"), results.workbookSummary().workbook().sheetNames());
      assertEquals("BudgetTotal", results.namedRanges().namedRanges().getFirst().name());
      assertEquals("Budget", results.sheetSummary().sheet().sheetName());
      assertEquals("A1", results.cells().cells().getFirst().address());
      assertEquals("A1", results.window().window().rows().getFirst().cells().getFirst().address());
      assertEquals("A1:B1", results.mergedRegions().mergedRegions().getFirst().range());
      assertEquals(
          "https://example.com/report",
          results.hyperlinks().hyperlinks().getFirst().hyperlink().target());
      assertEquals("Review", results.comments().comments().getFirst().comment().text());
      assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), results.layout().layout().pane());
      assertEquals(130, results.layout().layout().zoomPercent());
      assertEquals(
          ExcelPrintOrientation.LANDSCAPE,
          results.printLayout().printLayout().layout().orientation());
      assertEquals(1, results.formulaSurface().analysis().totalFormulaCellCount());
      assertEquals(
          "SUM(B2:B3)",
          results.formulaSurface().analysis().sheets().getFirst().formulas().getFirst().formula());
      assertEquals("Budget", results.schema().analysis().sheetName());
      assertEquals(4, results.schema().analysis().dataRowCount());
      assertEquals(1, results.namedRangeSurface().analysis().rangeBackedCount());
      assertEquals(0, results.namedRangeSurface().analysis().formulaBackedCount());
      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("D1:E3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")),
          results.autofilters().autofilters());
      assertEquals(1, results.tables().tables().size());
      assertEquals("Queue", results.tables().tables().getFirst().name());
      assertTrue(results.tables().tables().getFirst().hasAutofilter());
    }
  }

  @Test
  void selectsNamedRangesByExactSelectorsAndRejectsMissingSelectors() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-introspector-ranges-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      workbook.createSheet("Forecast");
      var workbookName = workbook.createName();
      workbookName.setNameName("BudgetRollup");
      workbookName.setRefersToFormula("SUM(Budget!$B$2:$B$3)");
      var sheetName = workbook.createName();
      sheetName.setNameName("LocalItem");
      sheetName.setSheetIndex(0);
      sheetName.setRefersToFormula("Budget!$A$1");
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();

      List<ExcelNamedRangeSnapshot> all =
          introspector.selectNamedRanges(workbook, new ExcelNamedRangeSelection.All());
      List<ExcelNamedRangeSnapshot> selected =
          introspector.selectNamedRanges(
              workbook,
              new ExcelNamedRangeSelection.Selected(
                  List.of(
                      new ExcelNamedRangeSelector.ByName("budgetrollup"),
                      new ExcelNamedRangeSelector.WorkbookScope("BudgetRollup"),
                      new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))));
      assertEquals(2, all.size());
      assertEquals(List.of(all.getFirst(), all.get(1)), selected);
      NamedRangeNotFoundException missing =
          assertThrows(
              NamedRangeNotFoundException.class,
              () ->
                  introspector.selectNamedRanges(
                      workbook,
                      new ExcelNamedRangeSelection.Selected(
                          List.of(new ExcelNamedRangeSelector.ByName("MissingRange")))));
      assertEquals("MissingRange", missing.name());

      NamedRangeNotFoundException missingWorkbookScope =
          assertThrows(
              NamedRangeNotFoundException.class,
              () ->
                  introspector.selectNamedRanges(
                      workbook,
                      new ExcelNamedRangeSelection.Selected(
                          List.of(new ExcelNamedRangeSelector.WorkbookScope("MissingWorkbook")))));
      assertEquals("MissingWorkbook", missingWorkbookScope.name());
      assertInstanceOf(ExcelNamedRangeScope.WorkbookScope.class, missingWorkbookScope.scope());

      NamedRangeNotFoundException missingSheetScope =
          assertThrows(
              NamedRangeNotFoundException.class,
              () ->
                  introspector.selectNamedRanges(
                      workbook,
                      new ExcelNamedRangeSelection.Selected(
                          List.of(
                              new ExcelNamedRangeSelector.SheetScope("SharedName", "Budget")))));
      assertEquals("SharedName", missingSheetScope.name());
      ExcelNamedRangeScope.SheetScope scope =
          assertInstanceOf(ExcelNamedRangeScope.SheetScope.class, missingSheetScope.scope());
      assertEquals("Budget", scope.sheetName());
    }
  }

  @Test
  void matchSelectorHandlesSameNameScopeEdgeCases() {
    List<ExcelNamedRangeSnapshot> namedRanges =
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "SharedName",
                new ExcelNamedRangeScope.WorkbookScope(),
                "Budget!$A$1",
                new ExcelNamedRangeTarget("Budget", "A1")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "SharedName",
                new ExcelNamedRangeScope.SheetScope("Forecast"),
                "Forecast!$A$1",
                new ExcelNamedRangeTarget("Forecast", "A1")));

    List<ExcelNamedRangeSnapshot> workbookScoped =
        ExcelWorkbookIntrospector.matchSelector(
            namedRanges, new ExcelNamedRangeSelector.WorkbookScope("SharedName"));
    List<ExcelNamedRangeSnapshot> sheetScoped =
        ExcelWorkbookIntrospector.matchSelector(
            namedRanges, new ExcelNamedRangeSelector.SheetScope("SharedName", "Forecast"));

    assertEquals(1, workbookScoped.size());
    assertInstanceOf(ExcelNamedRangeScope.WorkbookScope.class, workbookScoped.getFirst().scope());
    assertEquals(1, sheetScoped.size());
    ExcelNamedRangeScope.SheetScope scope =
        assertInstanceOf(ExcelNamedRangeScope.SheetScope.class, sheetScoped.getFirst().scope());
    assertEquals("Forecast", scope.sheetName());

    NamedRangeNotFoundException missingWorkbookSelector =
        assertThrows(
            NamedRangeNotFoundException.class,
            () ->
                ExcelWorkbookIntrospector.matchSelector(
                    namedRanges, new ExcelNamedRangeSelector.WorkbookScope("Missing")));
    assertEquals("Missing", missingWorkbookSelector.name());

    NamedRangeNotFoundException wrongSheetSelector =
        assertThrows(
            NamedRangeNotFoundException.class,
            () ->
                ExcelWorkbookIntrospector.matchSelector(
                    namedRanges, new ExcelNamedRangeSelector.SheetScope("SharedName", "Budget")));
    assertEquals("SharedName", wrongSheetSelector.name());
  }

  @Test
  void executesWorkbookProtectionIntrospectionAgainstWorkbookState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.xssfWorkbook().lockStructure();
      workbook.xssfWorkbook().lockRevision();
      workbook.xssfWorkbook().setWorkbookPassword("secret", HashAlgorithm.sha512);
      var protection =
          workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection()
              ? workbook.xssfWorkbook().getCTWorkbook().getWorkbookProtection()
              : workbook.xssfWorkbook().getCTWorkbook().addNewWorkbookProtection();
      protection.setWorkbookPassword(new byte[] {0x01, 0x02});

      WorkbookReadResult.WorkbookProtectionResult result =
          assertInstanceOf(
              WorkbookReadResult.WorkbookProtectionResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetWorkbookProtection("workbook-protection")));

      assertEquals("workbook-protection", result.stepId());
      assertTrue(result.protection().structureLocked());
      assertFalse(result.protection().windowsLocked());
      assertTrue(result.protection().revisionsLocked());
      assertTrue(result.protection().workbookPasswordHashPresent());
      assertFalse(result.protection().revisionsPasswordHashPresent());
    }
  }

  @Test
  void executesDrawingIntrospectionAgainstWorkbookState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .setPicture(
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(
                      java.util.Base64.getDecoder()
                          .decode(
                              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=")),
                  ExcelPictureFormat.PNG,
                  new ExcelDrawingAnchor.TwoCell(
                      new ExcelDrawingMarker(1, 2),
                      new ExcelDrawingMarker(4, 6),
                      ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                  "Queue preview"))
          .setEmbeddedObject(
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(
                      java.util.Base64.getDecoder()
                          .decode(
                              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=")),
                  new ExcelDrawingAnchor.TwoCell(
                      new ExcelDrawingMarker(5, 2),
                      new ExcelDrawingMarker(8, 6),
                      ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE)));

      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      WorkbookReadResult.DrawingObjectsResult drawingObjects =
          assertInstanceOf(
              WorkbookReadResult.DrawingObjectsResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetDrawingObjects("drawing", "Ops")));
      WorkbookReadResult.DrawingObjectPayloadResult drawingPayload =
          assertInstanceOf(
              WorkbookReadResult.DrawingObjectPayloadResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDrawingObjectPayload("payload", "Ops", "OpsEmbed")));

      assertEquals(2, drawingObjects.drawingObjects().size());
      assertEquals("OpsPicture", drawingObjects.drawingObjects().getFirst().name());
      assertEquals("OpsEmbed", drawingPayload.payload().name());
    }
  }

  @Test
  void rejectsNullWorkbookAndCommands() {
    ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();

    assertThrows(
        NullPointerException.class,
        () -> introspector.execute(null, new WorkbookReadCommand.GetWorkbookSummary("workbook")));
    assertThrows(
        NullPointerException.class, () -> introspector.execute(ExcelWorkbook.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> introspector.selectNamedRanges(null, new ExcelNamedRangeSelection.All()));
    assertThrows(
        NullPointerException.class,
        () -> introspector.selectNamedRanges(ExcelWorkbook.create(), null));
  }

  @Test
  void getFormulaSurfaceAndNamedRangeSurfaceRespectSelections() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-introspector-surface-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var budget = workbook.createSheet("Budget");
      budget.createRow(0).createCell(0).setCellValue("Item");
      budget.getRow(0).createCell(1).setCellValue("Amount");
      budget.createRow(1).createCell(0).setCellValue("Hosting");
      budget.getRow(1).createCell(1).setCellFormula("1+1");

      var forecast = workbook.createSheet("Forecast");
      forecast.createRow(0).createCell(0).setCellValue("Item");
      forecast.getRow(0).createCell(1).setCellValue("Amount");
      forecast.createRow(1).createCell(0).setCellValue("Domain");
      forecast.getRow(1).createCell(1).setCellFormula("2+2");

      var workbookName = workbook.createName();
      workbookName.setNameName("BudgetRollup");
      workbookName.setRefersToFormula("SUM(Budget!$B$2:$B$2)");
      var sheetName = workbook.createName();
      sheetName.setNameName("LocalItem");
      sheetName.setSheetIndex(0);
      sheetName.setRefersToFormula("Budget!$A$2");

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();

      WorkbookReadResult.FormulaSurfaceResult formulaSurface =
          cast(
              WorkbookReadResult.FormulaSurfaceResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetFormulaSurface(
                      "formula", new ExcelSheetSelection.Selected(List.of("Forecast")))));
      assertEquals(1, formulaSurface.analysis().totalFormulaCellCount());
      assertEquals(
          List.of("Forecast"),
          formulaSurface.analysis().sheets().stream()
              .map(WorkbookReadResult.SheetFormulaSurface::sheetName)
              .toList());

      WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurface =
          cast(
              WorkbookReadResult.NamedRangeSurfaceResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetNamedRangeSurface(
                      "namedRangeSurface",
                      new ExcelNamedRangeSelection.Selected(
                          List.of(
                              new ExcelNamedRangeSelector.WorkbookScope("BudgetRollup"),
                              new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))))));
      assertEquals(1, namedRangeSurface.analysis().workbookScopedCount());
      assertEquals(1, namedRangeSurface.analysis().sheetScopedCount());
      assertEquals(1, namedRangeSurface.analysis().formulaBackedCount());
      assertEquals(1, namedRangeSurface.analysis().rangeBackedCount());
    }
  }

  @Test
  void getSheetSchemaReturnsNullDominantTypeOnTies() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");
      sheet.setCell("A1", ExcelCellValue.text("Mixed"));
      sheet.setCell("A2", ExcelCellValue.text("text"));
      sheet.setCell("A3", ExcelCellValue.number(1.0));

      WorkbookReadResult.SheetSchemaResult schema =
          cast(
              WorkbookReadResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetSheetSchema("schema", "Budget", "A1", 3, 1)));

      assertNull(schema.analysis().columns().getFirst().dominantType());
    }
  }

  @Test
  void schemaDataRowCountIsZeroWhenHeaderRowIsEntirelyBlank() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      // empty sheet — every cell in the header row is blank
      workbook.getOrCreateSheet("Empty");

      WorkbookReadResult.SheetSchemaResult schema =
          cast(
              WorkbookReadResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetSheetSchema("schema", "Empty", "A1", 5, 3)));

      assertEquals(
          0,
          schema.analysis().dataRowCount(),
          "dataRowCount must be 0 when all header cells are blank");
      assertEquals(3, schema.analysis().columns().size());
    }
  }

  @Test
  void schemaDataRowCountIsRowCountMinusOneWhenHeaderIsPopulated() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Data");
      sheet.setCell("A1", ExcelCellValue.text("Name"));
      sheet.setCell("B1", ExcelCellValue.text("Score"));
      sheet.setCell("A2", ExcelCellValue.text("Alice"));
      sheet.setCell("B2", ExcelCellValue.number(95.0));

      WorkbookReadResult.SheetSchemaResult schema =
          cast(
              WorkbookReadResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetSheetSchema("schema", "Data", "A1", 3, 2)));

      assertEquals(2, schema.analysis().dataRowCount(), "dataRowCount must be rowCount - 1");
    }
  }

  @Test
  void schemaUsesEvaluatedTypeForFormulaCells() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Data");
      sheet.setCell("A1", ExcelCellValue.text("Total"));
      sheet.setCell("A2", ExcelCellValue.formula("1+1"));
      sheet.setCell("A3", ExcelCellValue.formula("2+2"));

      WorkbookReadResult.SheetSchemaResult schema =
          cast(
              WorkbookReadResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetSheetSchema("schema", "Data", "A1", 3, 1)));

      WorkbookReadResult.SchemaColumn column = schema.analysis().columns().getFirst();
      assertEquals(1, column.observedTypes().size());
      assertEquals("NUMBER", column.observedTypes().getFirst().type());
      assertEquals("NUMBER", column.dominantType());
    }
  }

  @Test
  void schemaCountsBooleanAndErrorCellTypes() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-introspector-types-", ".xlsx");
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      var poiSheet = poiWorkbook.createSheet("Data");
      poiSheet.createRow(0).createCell(0).setCellValue("Flag");
      poiSheet.getRow(0).createCell(1).setCellValue("Err");
      poiSheet.createRow(1).createCell(0).setCellValue(true);
      poiSheet
          .getRow(1)
          .createCell(1)
          .setCellErrorValue(org.apache.poi.ss.usermodel.FormulaError.DIV0.getCode());
      poiSheet.createRow(2).createCell(0).setCellValue(false);
      poiSheet
          .getRow(2)
          .createCell(1)
          .setCellErrorValue(org.apache.poi.ss.usermodel.FormulaError.DIV0.getCode());
      try (var out = Files.newOutputStream(workbookPath)) {
        poiWorkbook.write(out);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadResult.SheetSchemaResult schema =
          cast(
              WorkbookReadResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetSheetSchema("schema", "Data", "A1", 3, 2)));

      WorkbookReadResult.SchemaColumn boolCol = schema.analysis().columns().get(0);
      assertEquals("BOOLEAN", boolCol.dominantType());

      WorkbookReadResult.SchemaColumn errCol = schema.analysis().columns().get(1);
      assertEquals("ERROR", errCol.dominantType());
    }
  }

  @Test
  void schemaDominantTypeIsNullWhenMinorityTypeExistsButTieDoesNot() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Data");
      sheet.setCell("A1", ExcelCellValue.text("Col"));
      sheet.setCell("A2", ExcelCellValue.text("a"));
      sheet.setCell("A3", ExcelCellValue.text("b"));
      sheet.setCell("A4", ExcelCellValue.number(1.0));

      WorkbookReadResult.SheetSchemaResult schema =
          cast(
              WorkbookReadResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetSheetSchema("schema", "Data", "A1", 4, 1)));

      assertEquals("STRING", schema.analysis().columns().getFirst().dominantType());
    }
  }

  private void populateWorkbookForFullIntrospection(ExcelWorkbook workbook) {
    ExcelSheet budget = workbook.getOrCreateSheet("Budget");
    budget.setCell("A1", ExcelCellValue.text("Report"));
    budget.setCell("A2", ExcelCellValue.text("Hosting"));
    budget.setCell("B2", ExcelCellValue.number(49.0));
    budget.setCell("A3", ExcelCellValue.text("Domain"));
    budget.setCell("B3", ExcelCellValue.number(12.0));
    budget.setCell("B4", ExcelCellValue.number(61.0));
    budget.setCell("B5", ExcelCellValue.formula("SUM(B2:B3)"));
    budget.mergeCells("A1:B1");
    budget.setColumnWidth(0, 0, 12.5);
    budget.setRowHeight(0, 0, 18.0);
    budget.setPane(new ExcelSheetPane.Frozen(1, 1, 1, 1));
    budget.setZoom(130);
    budget.setPrintLayout(
        new ExcelPrintLayout(
            new ExcelPrintLayout.Area.Range("A1:B5"),
            ExcelPrintOrientation.LANDSCAPE,
            new ExcelPrintLayout.Scaling.Fit(1, 0),
            new ExcelPrintLayout.TitleRows.Band(0, 0),
            new ExcelPrintLayout.TitleColumns.None(),
            new ExcelHeaderFooterText("Budget", "", ""),
            new ExcelHeaderFooterText("", "Page &P", "")));
    budget.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
    budget.setComment("A1", new ExcelComment("Review", "GridGrind", false));
    workbook.setNamedRange(
        new ExcelNamedRangeDefinition(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            new ExcelNamedRangeTarget("Budget", "B4")));

    ExcelSheet ops = workbook.getOrCreateSheet("Ops");
    ops.setCell("A1", ExcelCellValue.text("Owner"));
    ops.setCell("B1", ExcelCellValue.text("Task"));
    ops.setCell("A2", ExcelCellValue.text("Ada"));
    ops.setCell("B2", ExcelCellValue.text("Queue"));
    ops.setCell("A3", ExcelCellValue.text("Lin"));
    ops.setCell("B3", ExcelCellValue.text("Pack"));
    ops.setCell("D1", ExcelCellValue.text("Region"));
    ops.setCell("E1", ExcelCellValue.text("Desk"));
    ops.setCell("D2", ExcelCellValue.text("North"));
    ops.setCell("E2", ExcelCellValue.text("A1"));
    ops.setCell("D3", ExcelCellValue.text("South"));
    ops.setCell("E3", ExcelCellValue.text("B1"));
    ops.setAutofilter("D1:E3");
    workbook.setTable(
        new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));
  }

  private IntrospectionReadResults readEveryIntrospectionResult(
      ExcelWorkbook workbook, ExcelWorkbookIntrospector introspector) {
    return new IntrospectionReadResults(
        cast(
            WorkbookReadResult.WorkbookSummaryResult.class,
            introspector.execute(workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook"))),
        cast(
            WorkbookReadResult.NamedRangesResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetNamedRanges(
                    "ranges", new ExcelNamedRangeSelection.All()))),
        cast(
            WorkbookReadResult.SheetSummaryResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetSheetSummary("sheet", "Budget"))),
        cast(
            WorkbookReadResult.CellsResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetCells("cells", "Budget", List.of("A1", "B4")))),
        cast(
            WorkbookReadResult.WindowResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetWindow("window", "Budget", "A1", 2, 2))),
        cast(
            WorkbookReadResult.MergedRegionsResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetMergedRegions("merged", "Budget"))),
        cast(
            WorkbookReadResult.HyperlinksResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetHyperlinks(
                    "hyperlinks", "Budget", new ExcelCellSelection.Selected(List.of("A1"))))),
        cast(
            WorkbookReadResult.CommentsResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetComments(
                    "comments", "Budget", new ExcelCellSelection.AllUsedCells()))),
        cast(
            WorkbookReadResult.SheetLayoutResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetSheetLayout("layout", "Budget"))),
        cast(
            WorkbookReadResult.PrintLayoutResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetPrintLayout("printLayout", "Budget"))),
        cast(
            WorkbookReadResult.FormulaSurfaceResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetFormulaSurface(
                    "formula", new ExcelSheetSelection.All()))),
        cast(
            WorkbookReadResult.SheetSchemaResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetSheetSchema("schema", "Budget", "A1", 5, 2))),
        cast(
            WorkbookReadResult.NamedRangeSurfaceResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetNamedRangeSurface(
                    "namedRangeSurface", new ExcelNamedRangeSelection.All()))),
        cast(
            WorkbookReadResult.AutofiltersResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetAutofilters("autofilters", "Ops"))),
        cast(
            WorkbookReadResult.TablesResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All()))));
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }

  private record IntrospectionReadResults(
      WorkbookReadResult.WorkbookSummaryResult workbookSummary,
      WorkbookReadResult.NamedRangesResult namedRanges,
      WorkbookReadResult.SheetSummaryResult sheetSummary,
      WorkbookReadResult.CellsResult cells,
      WorkbookReadResult.WindowResult window,
      WorkbookReadResult.MergedRegionsResult mergedRegions,
      WorkbookReadResult.HyperlinksResult hyperlinks,
      WorkbookReadResult.CommentsResult comments,
      WorkbookReadResult.SheetLayoutResult layout,
      WorkbookReadResult.PrintLayoutResult printLayout,
      WorkbookReadResult.FormulaSurfaceResult formulaSurface,
      WorkbookReadResult.SheetSchemaResult schema,
      WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurface,
      WorkbookReadResult.AutofiltersResult autofilters,
      WorkbookReadResult.TablesResult tables) {}
}
