package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for ExcelWorkbookIntrospector workbook-fact reads and named-range selection. */
class ExcelWorkbookIntrospectorTest {
  @Test
  void readsOpaqueRawFormulaFactsWithoutEvaluatorExecution() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Raw");
      sheet.cells().setCell("A1", new ResolvedRawFormula("_xlfn.UNIQUE({1,2,2})"));
      ExcelSheet schemaSheet = workbook.getOrCreateSheet("Schema");
      schemaSheet.cells().setCell("A1", new CellValue.Text("Count"));
      schemaSheet.cells().setCell("A2", new ResolvedFormula("1+1"));

      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      WorkbookSurfaceResult.FormulaSurfaceResult formulaSurface =
          assertInstanceOf(
              WorkbookSurfaceResult.FormulaSurfaceResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetFormulaSurface(
                      new StepId("formula-surface"),
                      new ExcelSheetSelection.Selected(List.of("Raw")))));
      WorkbookSheetResult.CellsResult cells =
          assertInstanceOf(
              WorkbookSheetResult.CellsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetCells(
                      new StepId("formula-cell"),
                      "Raw",
                      List.of("A1"),
                      new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.FORMULA)))));
      WorkbookSurfaceResult.SheetSchemaResult schema =
          assertInstanceOf(
              WorkbookSurfaceResult.SheetSchemaResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetSheetSchema(
                      new StepId("formula-schema"),
                      "Schema",
                      new ExcelReadWindow("A1", 2, 1),
                      new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.FORMULA)))));

      assertEquals(1, formulaSurface.surface().totalFormulaCellCount());
      assertEquals(
          "_xlfn.UNIQUE({1,2,2})",
          formulaSurface.surface().sheets().getFirst().formulas().getFirst().formula());
      ExcelCellSnapshot.FormulaSnapshot formulaCell =
          assertInstanceOf(ExcelCellSnapshot.FormulaSnapshot.class, cells.cells().getFirst());
      assertEquals("_xlfn.UNIQUE({1,2,2})", formulaCell.formula());
      assertTrue(formulaCell.evaluation().isEmpty());
      assertEquals(
          "NUMBER", schema.surface().columns().getFirst().observedTypes().getFirst().type());
    }
  }

  @Test
  void executesEveryIntrospectionCommandAgainstWorkbookState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
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
      assertEquals(1, results.formulaSurface().surface().totalFormulaCellCount());
      assertEquals(
          "SUM(B2:B3)",
          results.formulaSurface().surface().sheets().getFirst().formulas().getFirst().formula());
      assertEquals("Budget", results.schema().surface().sheetName());
      assertEquals(4, results.schema().surface().dataRowCount());
      assertEquals(1, results.namedRangeSurface().surface().rangeBackedCount());
      assertEquals(0, results.namedRangeSurface().surface().formulaBackedCount());
      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("D1:E3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")),
          results.autofilters().autofilters());
      assertEquals(1, results.tables().tables().size());
      assertEquals("Queue", results.tables().tables().getFirst().name());
      assertTrue(results.tables().tables().getFirst().behavior().hasAutofilter());
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

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
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
                ExcelNamedRangeTarget.range("Budget", "A1")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "SharedName",
                new ExcelNamedRangeScope.SheetScope("Forecast"),
                "Forecast!$A$1",
                ExcelNamedRangeTarget.range("Forecast", "A1")));

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
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.xssfWorkbook().lockStructure();
      workbook.xssfWorkbook().lockRevision();
      workbook.xssfWorkbook().setWorkbookPassword("secret", HashAlgorithm.sha512);
      var protection =
          workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection()
              ? workbook.xssfWorkbook().getCTWorkbook().getWorkbookProtection()
              : workbook.xssfWorkbook().getCTWorkbook().addNewWorkbookProtection();
      protection.setWorkbookPassword(new byte[] {0x01, 0x02});

      WorkbookCoreResult.WorkbookProtectionResult result =
          assertInstanceOf(
              WorkbookCoreResult.WorkbookProtectionResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetWorkbookProtection(
                          new StepId("workbook-protection"))));

      assertEquals("workbook-protection", result.stepId().value());
      assertTrue(result.protection().structureLocked());
      assertFalse(result.protection().windowsLocked());
      assertTrue(result.protection().revisionsLocked());
      assertTrue(result.protection().workbookPasswordHashPresent());
      assertFalse(result.protection().revisionsPasswordHashPresent());
    }
  }

  @Test
  void executesDrawingIntrospectionAgainstWorkbookState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet
          .drawings()
          .setPicture(
              new ResolvedPicture(
                  "OpsPicture",
                  new BinaryContent(
                      java.util.Base64.getDecoder()
                          .decode(
                              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC")),
                  ExcelPictureFormat.PNG,
                  new DrawingAnchorSpec.TwoCell(
                      new DrawingMarkerSpec(1, 2, 0, 0),
                      new DrawingMarkerSpec(4, 6, 0, 0),
                      ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                  java.util.Optional.of("Queue preview")))
          .setEmbeddedObject(
              new ResolvedEmbeddedObject(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new BinaryContent("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new BinaryContent(
                      java.util.Base64.getDecoder()
                          .decode(
                              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC")),
                  new DrawingAnchorSpec.TwoCell(
                      new DrawingMarkerSpec(5, 2, 0, 0),
                      new DrawingMarkerSpec(8, 6, 0, 0),
                      ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE)));

      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      WorkbookDrawingResult.DrawingObjectsResult drawingObjects =
          assertInstanceOf(
              WorkbookDrawingResult.DrawingObjectsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDrawingObjects(new StepId("drawing"), "Ops")));
      WorkbookDrawingResult.DrawingObjectPayloadResult drawingPayload =
          assertInstanceOf(
              WorkbookDrawingResult.DrawingObjectPayloadResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDrawingObjectPayload(
                      new StepId("payload"), "Ops", "OpsEmbed")));

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
        () ->
            introspector.execute(
                null, new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook"))));
    assertThrows(
        NullPointerException.class,
        () ->
            introspector.execute(
                ExcelWorkbooks.create(), (WorkbookReadCommand.Introspection) null));
    assertThrows(
        NullPointerException.class,
        () -> introspector.selectNamedRanges(null, new ExcelNamedRangeSelection.All()));
    assertThrows(
        NullPointerException.class,
        () -> introspector.selectNamedRanges(ExcelWorkbooks.create(), null));
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

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();

      WorkbookSurfaceResult.FormulaSurfaceResult formulaSurface =
          cast(
              WorkbookSurfaceResult.FormulaSurfaceResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetFormulaSurface(
                      new StepId("formula"),
                      new ExcelSheetSelection.Selected(List.of("Forecast")))));
      assertEquals(1, formulaSurface.surface().totalFormulaCellCount());
      assertEquals(
          List.of("Forecast"),
          formulaSurface.surface().sheets().stream()
              .map(WorkbookSurfaceResult.SheetFormulaSurface::sheetName)
              .toList());

      WorkbookSurfaceResult.NamedRangeSurfaceResult namedRangeSurface =
          cast(
              WorkbookSurfaceResult.NamedRangeSurfaceResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetNamedRangeSurface(
                      new StepId("namedRangeSurface"),
                      new ExcelNamedRangeSelection.Selected(
                          List.of(
                              new ExcelNamedRangeSelector.WorkbookScope("BudgetRollup"),
                              new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))))));
      assertEquals(1, namedRangeSurface.surface().workbookScopedCount());
      assertEquals(1, namedRangeSurface.surface().sheetScopedCount());
      assertEquals(1, namedRangeSurface.surface().formulaBackedCount());
      assertEquals(1, namedRangeSurface.surface().rangeBackedCount());
    }
  }

  @Test
  void getSheetSchemaReturnsNullDominantTypeOnTies() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Budget");
      sheet.cells().setCell("A1", new CellValue.Text("Mixed"));
      sheet.cells().setCell("A2", new CellValue.Text("text"));
      sheet.cells().setCell("A3", new CellValue.NumberValue(1.0));

      WorkbookSurfaceResult.SheetSchemaResult schema =
          cast(
              WorkbookSurfaceResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      WorkbookReadTestSupport.getSheetSchema("schema", "Budget", "A1", 3, 1)));

      assertNull(schema.surface().columns().getFirst().dominantType());
    }
  }

  @Test
  void schemaDataRowCountIsZeroWhenHeaderRowIsEntirelyBlank() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      // empty sheet — every cell in the header row is blank
      workbook.getOrCreateSheet("Empty");

      WorkbookSurfaceResult.SheetSchemaResult schema =
          cast(
              WorkbookSurfaceResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      WorkbookReadTestSupport.getSheetSchema("schema", "Empty", "A1", 5, 3)));

      assertEquals(
          0,
          schema.surface().dataRowCount(),
          "dataRowCount must be 0 when all header cells are blank");
      assertEquals(3, schema.surface().columns().size());
    }
  }

  @Test
  void schemaDataRowCountIsRowCountMinusOneWhenHeaderIsPopulated() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Data");
      sheet.cells().setCell("A1", new CellValue.Text("Name"));
      sheet.cells().setCell("B1", new CellValue.Text("Score"));
      sheet.cells().setCell("A2", new CellValue.Text("Alice"));
      sheet.cells().setCell("B2", new CellValue.NumberValue(95.0));

      WorkbookSurfaceResult.SheetSchemaResult schema =
          cast(
              WorkbookSurfaceResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      WorkbookReadTestSupport.getSheetSchema("schema", "Data", "A1", 3, 2)));

      assertEquals(2, schema.surface().dataRowCount(), "dataRowCount must be rowCount - 1");
    }
  }

  @Test
  void schemaUsesEvaluatedTypeForFormulaCells() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Data");
      sheet.cells().setCell("A1", new CellValue.Text("Total"));
      sheet.cells().setCell("A2", new ResolvedFormula("1+1"));
      sheet.cells().setCell("A3", new ResolvedFormula("2+2"));

      WorkbookSurfaceResult.SheetSchemaResult schema =
          cast(
              WorkbookSurfaceResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      WorkbookReadTestSupport.getSheetSchema("schema", "Data", "A1", 3, 1)));

      WorkbookSurfaceResult.SchemaColumn column = schema.surface().columns().getFirst();
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

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      WorkbookSurfaceResult.SheetSchemaResult schema =
          cast(
              WorkbookSurfaceResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      WorkbookReadTestSupport.getSheetSchema("schema", "Data", "A1", 3, 2)));

      WorkbookSurfaceResult.SchemaColumn boolCol = schema.surface().columns().get(0);
      assertEquals("BOOLEAN", boolCol.dominantType());

      WorkbookSurfaceResult.SchemaColumn errCol = schema.surface().columns().get(1);
      assertEquals("ERROR", errCol.dominantType());
    }
  }

  @Test
  void schemaDominantTypeIsNullWhenMinorityTypeExistsButTieDoesNot() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Data");
      sheet.cells().setCell("A1", new CellValue.Text("Col"));
      sheet.cells().setCell("A2", new CellValue.Text("a"));
      sheet.cells().setCell("A3", new CellValue.Text("b"));
      sheet.cells().setCell("A4", new CellValue.NumberValue(1.0));

      WorkbookSurfaceResult.SheetSchemaResult schema =
          cast(
              WorkbookSurfaceResult.SheetSchemaResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      WorkbookReadTestSupport.getSheetSchema("schema", "Data", "A1", 4, 1)));

      assertEquals("TEXT", schema.surface().columns().getFirst().dominantType());
    }
  }

  private void populateWorkbookForFullIntrospection(ExcelWorkbook workbook) {
    ExcelSheet budget = workbook.getOrCreateSheet("Budget");
    budget.cells().setCell("A1", new CellValue.Text("Report"));
    budget.cells().setCell("A2", new CellValue.Text("Hosting"));
    budget.cells().setCell("B2", new CellValue.NumberValue(49.0));
    budget.cells().setCell("A3", new CellValue.Text("Domain"));
    budget.cells().setCell("B3", new CellValue.NumberValue(12.0));
    budget.cells().setCell("B4", new CellValue.NumberValue(61.0));
    budget.cells().setCell("B5", new ResolvedFormula("SUM(B2:B3)"));
    budget.layout().mergeCells("A1:B1");
    budget.columns().setWidth(0, 0, 12.5);
    budget.rows().setHeight(0, 0, 18.0);
    budget.layout().setPane(new ExcelSheetPane.Frozen(1, 1, 1, 1));
    budget.layout().setZoom(130);
    budget
        .layout()
        .setPrintLayout(
            new ExcelPrintLayout(
                new ExcelPrintLayout.Area.Range("A1:B5"),
                ExcelPrintOrientation.LANDSCAPE,
                new ExcelPrintLayout.Scaling.Fit(1, 0),
                new ExcelPrintLayout.TitleRows.Band(0, 0),
                new ExcelPrintLayout.TitleColumns.None(),
                new ExcelHeaderFooterText("Budget", "", ""),
                new ExcelHeaderFooterText("", "Page &P", "")));
    budget.annotations().setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
    budget.annotations().setComment("A1", new ResolvedComment("Review", "GridGrind", false));
    workbook
        .names()
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Budget", "B4")));

    ExcelSheet ops = workbook.getOrCreateSheet("Ops");
    ops.cells().setCell("A1", new CellValue.Text("Owner"));
    ops.cells().setCell("B1", new CellValue.Text("Task"));
    ops.cells().setCell("A2", new CellValue.Text("Ada"));
    ops.cells().setCell("B2", new CellValue.Text("Queue"));
    ops.cells().setCell("A3", new CellValue.Text("Lin"));
    ops.cells().setCell("B3", new CellValue.Text("Pack"));
    ops.cells().setCell("D1", new CellValue.Text("Region"));
    ops.cells().setCell("E1", new CellValue.Text("Desk"));
    ops.cells().setCell("D2", new CellValue.Text("North"));
    ops.cells().setCell("E2", new CellValue.Text("A1"));
    ops.cells().setCell("D3", new CellValue.Text("South"));
    ops.cells().setCell("E3", new CellValue.Text("B1"));
    ops.metadata().setAutofilter("D1:E3");
    workbook
        .tables()
        .setTable(
            new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));
  }

  private IntrospectionReadResults readEveryIntrospectionResult(
      ExcelWorkbook workbook, ExcelWorkbookIntrospector introspector) {
    return new IntrospectionReadResults(
        cast(
            WorkbookCoreResult.WorkbookSummaryResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook")))),
        cast(
            WorkbookCoreResult.NamedRangesResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetNamedRanges(
                    new StepId("ranges"), new ExcelNamedRangeSelection.All()))),
        cast(
            WorkbookSheetResult.SheetSummaryResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetSheetSummary(new StepId("sheet"), "Budget"))),
        cast(
            WorkbookSheetResult.CellsResult.class,
            introspector.execute(
                workbook,
                WorkbookReadTestSupport.getCells("cells", "Budget", List.of("A1", "B4")))),
        cast(
            WorkbookSheetResult.WindowResult.class,
            introspector.execute(
                workbook, WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 2, 2))),
        cast(
            WorkbookSheetResult.MergedRegionsResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetMergedRegions(new StepId("merged"), "Budget"))),
        cast(
            WorkbookSheetResult.HyperlinksResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetHyperlinks(
                    new StepId("hyperlinks"),
                    "Budget",
                    new ExcelCellSelection.Selected(List.of("A1"))))),
        cast(
            WorkbookSheetResult.CommentsResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetComments(
                    new StepId("comments"), "Budget", new ExcelCellSelection.AllUsedCells()))),
        cast(
            WorkbookSheetResult.SheetLayoutResult.class,
            introspector.execute(
                workbook, new WorkbookReadCommand.GetSheetLayout(new StepId("layout"), "Budget"))),
        cast(
            WorkbookSheetResult.PrintLayoutResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetPrintLayout(new StepId("printLayout"), "Budget"))),
        cast(
            WorkbookSurfaceResult.FormulaSurfaceResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetFormulaSurface(
                    new StepId("formula"), new ExcelSheetSelection.All()))),
        cast(
            WorkbookSurfaceResult.SheetSchemaResult.class,
            introspector.execute(
                workbook, WorkbookReadTestSupport.getSheetSchema("schema", "Budget", "A1", 5, 2))),
        cast(
            WorkbookSurfaceResult.NamedRangeSurfaceResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetNamedRangeSurface(
                    new StepId("namedRangeSurface"), new ExcelNamedRangeSelection.All()))),
        cast(
            WorkbookRuleResult.AutofiltersResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetAutofilters(new StepId("autofilters"), "Ops"))),
        cast(
            WorkbookRuleResult.TablesResult.class,
            introspector.execute(
                workbook,
                new WorkbookReadCommand.GetTables(
                    new StepId("tables"), new ExcelTableSelection.All()))));
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }

  private record IntrospectionReadResults(
      WorkbookCoreResult.WorkbookSummaryResult workbookSummary,
      WorkbookCoreResult.NamedRangesResult namedRanges,
      WorkbookSheetResult.SheetSummaryResult sheetSummary,
      WorkbookSheetResult.CellsResult cells,
      WorkbookSheetResult.WindowResult window,
      WorkbookSheetResult.MergedRegionsResult mergedRegions,
      WorkbookSheetResult.HyperlinksResult hyperlinks,
      WorkbookSheetResult.CommentsResult comments,
      WorkbookSheetResult.SheetLayoutResult layout,
      WorkbookSheetResult.PrintLayoutResult printLayout,
      WorkbookSurfaceResult.FormulaSurfaceResult formulaSurface,
      WorkbookSurfaceResult.SheetSchemaResult schema,
      WorkbookSurfaceResult.NamedRangeSurfaceResult namedRangeSurface,
      WorkbookRuleResult.AutofiltersResult autofilters,
      WorkbookRuleResult.TablesResult tables) {}
}
