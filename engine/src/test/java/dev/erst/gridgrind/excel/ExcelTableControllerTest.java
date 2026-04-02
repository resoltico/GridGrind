package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests for direct table authoring, introspection, and health analysis. */
class ExcelTableControllerTest {
  private final ExcelTableController controller = new ExcelTableController();

  @Test
  void setTable_createsTableOwnedAutofilterAndClearsOverlappingSheetAutofilter() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      sheet.setAutofilter("A1:B3");

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      List<ExcelTableSnapshot> tables = controller.tables(workbook, new ExcelTableSelection.All());
      List<ExcelAutofilterSnapshot> autofilters = controller.tableOwnedAutofilters(workbook, "Ops");

      assertEquals(1, tables.size());
      assertEquals("Queue", tables.getFirst().name());
      assertTrue(tables.getFirst().hasAutofilter());
      assertFalse(workbook.sheet("Ops").xssfSheet().getCTWorksheet().isSetAutoFilter());
      assertEquals(List.of(new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")), autofilters);
    }
  }

  @Test
  void setTable_replacesSameNameOnSameSheetAndSupportsNamedStylesAndTotalsRow() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      sheet.setCell("A4", ExcelCellValue.text("Totals"));
      sheet.setCell("B4", ExcelCellValue.text("Done"));

      controller.setTable(
          workbook,
          new ExcelTableDefinition(
              "Queue",
              "Ops",
              "A1:B4",
              true,
              new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false)));
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      List<ExcelTableSnapshot> tables = controller.tables(workbook, new ExcelTableSelection.All());

      assertEquals(1, tables.size());
      assertEquals("A1:B3", tables.getFirst().range());
      assertEquals(0, tables.getFirst().totalsRowCount());
      assertInstanceOf(ExcelTableStyleSnapshot.None.class, tables.getFirst().style());
      assertEquals(
          List.of(new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")),
          controller.tableOwnedAutofilters(workbook, "Ops"));
    }
  }

  @Test
  void setTable_acceptsFormulaAndNumericHeaders() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setCell("A1", ExcelCellValue.formula("A2"));
      sheet.setCell("B1", ExcelCellValue.number(12.0));
      sheet.setCell("A2", ExcelCellValue.text("Ada"));
      sheet.setCell("B2", ExcelCellValue.text("Queue"));
      sheet.setCell("A3", ExcelCellValue.text("Lin"));
      sheet.setCell("B3", ExcelCellValue.text("Pack"));

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      ExcelTableSnapshot table =
          controller.tables(workbook, new ExcelTableSelection.All()).getFirst();
      assertEquals(2, table.columnNames().size());
      assertTrue(table.columnNames().stream().allMatch(name -> !name.isBlank()));
      assertEquals(
          List.of(), controller.tableHealthFindings(workbook, new ExcelTableSelection.All()));
    }
  }

  @Test
  void tables_selectByNamesPreservesRequestOrderAndIgnoresMissingNames() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet ops = workbook.getOrCreateSheet("Ops");
      ExcelSheet archive = workbook.getOrCreateSheet("Archive");
      populateTableCells(ops, "Owner", "Task");
      populateTableCells(archive, "Owner", "Task");

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));
      controller.setTable(
          workbook,
          new ExcelTableDefinition(
              "ArchiveQueue", "Archive", "A1:B3", false, new ExcelTableStyle.None()));

      List<ExcelTableSnapshot> selected =
          controller.tables(
              workbook,
              new ExcelTableSelection.ByNames(List.of("archivequeue", "missing", "QUEUE")));

      assertEquals(
          List.of("ArchiveQueue", "Queue"),
          selected.stream().map(ExcelTableSnapshot::name).toList());
      assertEquals(
          List.of("Archive", "Ops"), selected.stream().map(ExcelTableSnapshot::sheetName).toList());
    }
  }

  @Test
  void setTable_rejectsBlankHeaders() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "", "Task");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setTable(
                      workbook,
                      new ExcelTableDefinition(
                          "Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None())));

      assertEquals("table header cells must not be blank", failure.getMessage());
    }
  }

  @Test
  void setTable_rejectsCaseInsensitiveDuplicateHeaders() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "owner");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setTable(
                      workbook,
                      new ExcelTableDefinition(
                          "Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None())));

      assertEquals(
          "table header cells must be unique (case-insensitive): owner", failure.getMessage());
    }
  }

  @Test
  void setTable_rejectsOverlapAndDefinedNameConflicts() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      sheet.setCell("C1", ExcelCellValue.text("Desk"));
      sheet.setCell("C2", ExcelCellValue.text("A1"));
      sheet.setCell("C3", ExcelCellValue.text("B1"));
      sheet.setCell("D1", ExcelCellValue.text("Region"));
      sheet.setCell("E1", ExcelCellValue.text("Desk"));
      sheet.setCell("D2", ExcelCellValue.text("North"));
      sheet.setCell("E2", ExcelCellValue.text("A1"));
      sheet.setCell("D3", ExcelCellValue.text("South"));
      sheet.setCell("E3", ExcelCellValue.text("B1"));

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      IllegalArgumentException overlapFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setTable(
                      workbook,
                      new ExcelTableDefinition(
                          "RegionDesk", "Ops", "B1:C3", false, new ExcelTableStyle.None())));

      assertEquals(
          "table range must not overlap an existing table: Queue@A1:B3",
          overlapFailure.getMessage());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "Queue",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Ops", "A1")));

      IllegalArgumentException definedNameFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setTable(
                      workbook,
                      new ExcelTableDefinition(
                          "Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None())));

      assertEquals(
          "table name must not conflict with an existing defined name: Queue",
          definedNameFailure.getMessage());
    }
  }

  @Test
  void setTable_rejectsUnsupportedShapesUnknownStylesAndCrossSheetNameReuse() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet ops = workbook.getOrCreateSheet("Ops");
      populateTableCells(ops, "Owner", "Task");
      ExcelSheet archive = workbook.getOrCreateSheet("Archive");
      populateTableCells(archive, "Owner", "Task");

      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setTable(
                  workbook,
                  new ExcelTableDefinition(
                      "Queue", "Ops", "A1:B1", false, new ExcelTableStyle.None())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setTable(
                  workbook,
                  new ExcelTableDefinition(
                      "Queue", "Ops", "A1:B2", true, new ExcelTableStyle.None())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setTable(
                  workbook,
                  new ExcelTableDefinition(
                      "Queue",
                      "Ops",
                      "A1:B3",
                      false,
                      new ExcelTableStyle.Named("MissingStyle", false, false, true, false))));

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      IllegalArgumentException crossSheetFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setTable(
                      workbook,
                      new ExcelTableDefinition(
                          "Queue", "Archive", "A1:B3", false, new ExcelTableStyle.None())));
      assertEquals(
          "table name already exists on a different sheet: Queue", crossSheetFailure.getMessage());
    }
  }

  @Test
  void deleteTable_rejectsMissingExpectedSheet() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> controller.deleteTable(workbook, "Queue", "Archive"));
      assertEquals("table not found on expected sheet: Queue@Archive", failure.getMessage());
    }
  }

  @Test
  void deleteTable_removesExistingTableByWorkbookGlobalName() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      controller.deleteTable(workbook, "Queue", "Ops");

      assertEquals(List.of(), controller.tables(workbook, new ExcelTableSelection.All()));
      assertEquals(List.of(), controller.tableOwnedAutofilters(workbook, "Ops"));
    }
  }

  @Test
  void deleteTable_rejectsBlankExpectedSheet() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.deleteTable(workbook, "Queue", " "));

      assertEquals("sheetName must not be blank", failure.getMessage());
    }
  }

  @Test
  void tableHealthFindings_detectMalformedLoadedTables() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook, poiWorkbook.getCreationHelper().createFormulaEvaluator())) {
      XSSFSheet sheet = poiWorkbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Owner");
      sheet.getRow(0).createCell(1).setCellValue("Owner");
      sheet.getRow(0).createCell(2).setBlank();
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");
      sheet.getRow(1).createCell(2).setCellValue("North");
      sheet.createRow(2).createCell(0).setCellValue("Lin");
      sheet.getRow(2).createCell(1).setCellValue("Pack");
      sheet.getRow(2).createCell(2).setCellValue("South");

      XSSFTable first = sheet.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));
      XSSFTable second =
          sheet.createTable(new AreaReference("B1:C3", SpreadsheetVersion.EXCEL2007));
      first.getCTTable().addNewTableStyleInfo().setName("MissingStyle");
      second.getCTTable().addNewTableStyleInfo().setName("MissingStyle");

      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          controller.tableHealthFindings(workbook, new ExcelTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();

      assertTrue(codes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_DUPLICATE_HEADER));
      assertTrue(codes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_BLANK_HEADER));
      assertTrue(codes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_OVERLAPPING_RANGE));
      assertTrue(codes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_STYLE_MISMATCH));
    }
  }

  @Test
  void tableHealthAndAutofilterHealthReturnEmptyForHealthyStyledTotalsTable() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      sheet.setCell("A4", ExcelCellValue.text("Totals"));
      sheet.setCell("B4", ExcelCellValue.text("Done"));

      controller.setTable(
          workbook,
          new ExcelTableDefinition(
              "Queue",
              "Ops",
              "A1:B4",
              true,
              new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false)));

      ExcelTableSnapshot table =
          controller.tables(workbook, new ExcelTableSelection.All()).getFirst();
      assertEquals("A1:B3", controller.tableOwnedAutofilters(workbook, "Ops").getFirst().range());
      assertEquals("A1:B3", ExcelTableStructureSupport.expectedAutofilterRangeText(table));
      assertEquals(
          List.of(), controller.tableHealthFindings(workbook, new ExcelTableSelection.All()));
      assertEquals(List.of(), controller.tableAutofilterHealthFindings(workbook, "Ops"));
    }
  }

  @Test
  void setTable_keepsNonOverlappingSheetAutofilter() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      sheet.setCell("D1", ExcelCellValue.text("Desk"));
      sheet.setCell("D2", ExcelCellValue.text("A1"));
      sheet.setCell("D3", ExcelCellValue.text("B1"));
      sheet.setAutofilter("D1:D3");

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      assertEquals(
          List.of(new ExcelAutofilterSnapshot.SheetOwned("D1:D3")),
          new ExcelAutofilterController().sheetOwnedAutofilters(workbook.sheet("Ops").xssfSheet()));
    }
  }

  @Test
  void setTable_ignoresInvalidSheetAutofilterMetadata() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      workbook.sheet("Ops").xssfSheet().getCTWorksheet().addNewAutoFilter().setRef("A0:B3");

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      assertEquals(
          "A0:B3", workbook.sheet("Ops").xssfSheet().getCTWorksheet().getAutoFilter().getRef());
    }
  }

  @Test
  void tableHealthFindings_flagsShortRangesAndBlankStyleNames() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook, poiWorkbook.getCreationHelper().createFormulaEvaluator())) {
      XSSFSheet sheet = poiWorkbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Owner");
      sheet.getRow(0).createCell(1).setCellValue("Task");
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");

      XSSFTable table = sheet.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
      table.getCTTable().setTotalsRowCount(1);
      table.getCTTable().setTotalsRowShown(true);
      table.getCTTable().addNewTableStyleInfo().setName("");

      List<WorkbookAnalysis.AnalysisFindingCode> codes =
          controller.tableHealthFindings(workbook, new ExcelTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();

      assertTrue(codes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_BROKEN_REFERENCE));
      assertTrue(codes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_STYLE_MISMATCH));
    }
  }

  @Test
  void tableHealthFindings_flagsBrokenRangeMetadata() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      XSSFTable table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      table.getCTTable().setRef("A0:B3");

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.tableHealthFindings(workbook, new ExcelTableSelection.All());

      assertEquals(1, findings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.TABLE_BROKEN_REFERENCE, findings.getFirst().code());
    }
  }

  @Test
  void tableAutofilterHealthFindings_flagsMismatchedTableFilterRanges() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      XSSFTable table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      table.getCTTable().getAutoFilter().setRef("A1:B2");

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.tableAutofilterHealthFindings(workbook, "Ops");

      assertEquals(1, findings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          findings.getFirst().code());
    }
  }

  @Test
  void tableAutofilterHealthFindings_flagsInvalidRangesAndBlankHeaders() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      XSSFTable table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      table.getCTTable().getAutoFilter().setRef("A0:B3");

      List<WorkbookAnalysis.AnalysisFinding> invalidRangeFindings =
          controller.tableAutofilterHealthFindings(workbook, "Ops");

      assertEquals(1, invalidRangeFindings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
          invalidRangeFindings.getFirst().code());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));
      sheet.setCell("A1", ExcelCellValue.blank());
      sheet.setCell("B1", ExcelCellValue.blank());

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.tableAutofilterHealthFindings(workbook, "Ops");

      assertEquals(1, findings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
          findings.getFirst().code());
    }
  }

  @Test
  void tableAutofilterHealthFindings_usesSheetLocationWhenTableRangeMetadataIsInvalid()
      throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      XSSFTable table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      table.getCTTable().setRef("A0:B3");
      table.getCTTable().getAutoFilter().setRef("A1:B3");

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.tableAutofilterHealthFindings(workbook, "Ops");

      assertEquals(1, findings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          findings.getFirst().code());
      assertInstanceOf(
          WorkbookAnalysis.AnalysisLocation.Sheet.class, findings.getFirst().location());
    }
  }

  @Test
  void headerTextNormalizesSupportedCellKinds() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      var row = sheet.createRow(0);
      row.createCell(0).setBlank();
      row.createCell(1).setCellValue(" Owner ");
      var blankFormulaCell = row.createCell(2);
      blankFormulaCell.setCellFormula("A1");
      sheet.getCTWorksheet().getSheetData().getRowArray(0).getCArray(2).getF().setStringValue("");
      row.createCell(3).setCellFormula(" A1 ");
      row.createCell(4).setCellValue(12.0);
      row.createCell(5).setCellValue(true);

      assertEquals("", ExcelTableStructureSupport.headerText(null));
      assertEquals("", ExcelTableStructureSupport.headerText(sheet.getRow(0).getCell(0)));
      assertEquals("Owner", ExcelTableStructureSupport.headerText(sheet.getRow(0).getCell(1)));
      assertEquals("", ExcelTableStructureSupport.headerText(sheet.getRow(0).getCell(2)));
      assertEquals("A1", ExcelTableStructureSupport.headerText(sheet.getRow(0).getCell(3)));
      assertEquals("12.0", ExcelTableStructureSupport.headerText(sheet.getRow(0).getCell(4)));
      assertEquals("TRUE", ExcelTableStructureSupport.headerText(sheet.getRow(0).getCell(5)));
    }
  }

  @Test
  void applyStyleClearsExistingStyleInfoWhenNoneRequested() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Owner");
      sheet.getRow(0).createCell(1).setCellValue("Task");
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");
      XSSFTable table = sheet.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
      table.getCTTable().addNewTableStyleInfo().setName("TableStyleMedium2");

      ExcelTableStructureSupport.applyStyle(table, new ExcelTableStyle.None());

      assertFalse(table.getCTTable().isSetTableStyleInfo());
    }
  }

  @Test
  void requiredTableByNameRejectsMissingTable() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelTableCatalogSupport.requiredTableByName(sheet, "Queue"));

      assertEquals("table not found on sheet: Queue@Ops", failure.getMessage());
    }
  }

  @Test
  void tableOwnedAutofiltersReadsEmptyWhenSheetHasNoOwnedFilters() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet ops = workbook.getOrCreateSheet("Ops");
      populateTableCells(ops, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));
      workbook.sheet("Ops").xssfSheet().getTables().getFirst().getCTTable().unsetAutoFilter();

      assertEquals(List.of(), controller.tableOwnedAutofilters(workbook, "Ops"));
    }
  }

  @Test
  void tableAutofilterHealthFindingsReturnsEmptyWhenSheetHasNoRelevantTables() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet ops = workbook.getOrCreateSheet("Ops");
      populateTableCells(ops, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));
      workbook.sheet("Ops").xssfSheet().getTables().getFirst().getCTTable().unsetAutoFilter();

      assertEquals(List.of(), controller.tableAutofilterHealthFindings(workbook, "Ops"));
    }
  }

  @Test
  void setTable_ignoresNullNamedRangeNamesAndInvalidExistingTableRanges() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");

      org.apache.poi.ss.usermodel.Name unnamed = workbook.xssfWorkbook().createName();
      unnamed.setRefersToFormula("Ops!$D$1");
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "OtherAnchor",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Ops", "I1")));

      XSSFTable invalidTable =
          workbook
              .sheet("Ops")
              .xssfSheet()
              .createTable(new AreaReference("D1:E3", SpreadsheetVersion.EXCEL2007));
      invalidTable.setName("Legacy");
      invalidTable.setDisplayName("Legacy");
      invalidTable.getCTTable().setRef("D0:E3");

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      assertEquals(
          List.of("Legacy", "Queue"),
          controller.tables(workbook, new ExcelTableSelection.All()).stream()
              .map(ExcelTableSnapshot::name)
              .toList());
    }
  }

  @Test
  void setTable_replacesSameNameWhileIgnoringSelfForOverlapChecks() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      sheet.setCell("C1", ExcelCellValue.text("Desk"));
      sheet.setCell("C2", ExcelCellValue.text("A1"));
      sheet.setCell("C3", ExcelCellValue.text("B1"));
      sheet.setCell("E1", ExcelCellValue.text("Region"));
      sheet.setCell("F1", ExcelCellValue.text("Lead"));
      sheet.setCell("E2", ExcelCellValue.text("North"));
      sheet.setCell("F2", ExcelCellValue.text("Ada"));
      sheet.setCell("E3", ExcelCellValue.text("South"));
      sheet.setCell("F3", ExcelCellValue.text("Lin"));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "OpsAnchor",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Ops", "H1")));
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));
      controller.setTable(
          workbook,
          new ExcelTableDefinition(
              "RegionLead", "Ops", "E1:F3", false, new ExcelTableStyle.None()));

      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:C3", false, new ExcelTableStyle.None()));

      List<ExcelTableSnapshot> tables = controller.tables(workbook, new ExcelTableSelection.All());
      assertEquals(2, tables.size());
      assertTrue(
          tables.stream()
              .map(ExcelTableSnapshot::name)
              .toList()
              .containsAll(List.of("Queue", "RegionLead")));
      assertEquals(
          "A1:C3",
          tables.stream()
              .filter(table -> "Queue".equals(table.name()))
              .findFirst()
              .orElseThrow()
              .range());
    }
  }

  @Test
  void headerCellWritesSynchronizeTableMetadataImmediatelyAndAcrossSave(@TempDir Path tempDirectory)
      throws Exception {
    Path workbookPath = tempDirectory.resolve("table-header-cell-sync.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "c]cc", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition(
              "BudgetTable", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      XSSFTable table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      sheet.setCell("A1", ExcelCellValue.text("QQQQq"));

      assertEquals("QQQQq", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
      assertEquals(
          List.of("QQQQq", "Task"),
          controller.tables(workbook, new ExcelTableSelection.All()).getFirst().columnNames());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertEquals(
          List.of("QQQQq", "Task"),
          controller.tables(reopened, new ExcelTableSelection.All()).getFirst().columnNames());
    }
  }

  @Test
  void headerRangeWritesAndClearsSynchronizeTableMetadataAndHealthFindings(
      @TempDir Path tempDirectory) throws Exception {
    Path workbookPath = tempDirectory.resolve("table-header-range-sync.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      XSSFTable table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      sheet.setRange(
          "A1:B2",
          List.of(
              List.of(ExcelCellValue.text("Desk"), ExcelCellValue.text("Lane")),
              List.of(ExcelCellValue.text("Ada"), ExcelCellValue.text("Queue"))));

      assertEquals("Desk", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
      assertEquals("Lane", table.getCTTable().getTableColumns().getTableColumnArray(1).getName());

      sheet.clearRange("A1");

      assertEquals("", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
      assertTrue(
          controller.tableHealthFindings(workbook, new ExcelTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_BLANK_HEADER));

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      ExcelTableSnapshot table =
          controller.tables(reopened, new ExcelTableSelection.All()).getFirst();
      assertEquals(List.of("", "Lane"), table.columnNames());
      assertTrue(
          controller.tableHealthFindings(reopened, new ExcelTableSelection.All()).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_BLANK_HEADER));
    }
  }

  @Test
  void headerStyleWritesSynchronizeTableMetadataImmediatelyAndAcrossSave(
      @TempDir Path tempDirectory) throws Exception {
    Path workbookPath = tempDirectory.resolve("table-header-style-sync.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("V");
      sheet.appendRow(
          ExcelCellValue.dateTime(LocalDateTime.of(2026, 2, 6, 0, 19, 17)),
          ExcelCellValue.date(LocalDate.of(2026, 6, 26)));
      sheet.setCell("A2", ExcelCellValue.text("Ada"));
      sheet.setCell("B2", ExcelCellValue.text("Queue"));
      sheet.setCell("A3", ExcelCellValue.text("Totals"));
      sheet.setCell("B3", ExcelCellValue.text("Done"));
      controller.setTable(
          workbook,
          new ExcelTableDefinition(
              "OpsTable",
              "V",
              "A1:B3",
              true,
              new ExcelTableStyle.Named("TableStyleMedium2", true, true, true, true)));

      XSSFTable table = workbook.sheet("V").xssfSheet().getTables().getFirst();
      assertEquals(
          "2026-02-06 00:19:17",
          table.getCTTable().getTableColumns().getTableColumnArray(0).getName());

      sheet.applyStyle("A1:B2", ExcelCellStyle.numberFormat("yyyy-mm-dd"));

      assertEquals(
          "2026-02-06", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
      assertEquals(
          List.of("2026-02-06", "2026-06-26"),
          controller.tables(workbook, new ExcelTableSelection.All()).getFirst().columnNames());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertEquals(
          List.of("2026-02-06", "2026-06-26"),
          controller.tables(reopened, new ExcelTableSelection.All()).getFirst().columnNames());
    }
  }

  @Test
  void saveNormalizesStaleTypedHeaderMetadataBeforePersistence(@TempDir Path tempDirectory)
      throws Exception {
    Path workbookPath = tempDirectory.resolve("table-header-save-normalization.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("V");
      sheet.appendRow(
          ExcelCellValue.dateTime(LocalDateTime.of(2026, 2, 6, 0, 19, 17)),
          ExcelCellValue.date(LocalDate.of(2026, 6, 26)));
      sheet.setCell("A2", ExcelCellValue.text("Ada"));
      sheet.setCell("B2", ExcelCellValue.text("Queue"));
      sheet.setCell("A3", ExcelCellValue.text("Totals"));
      sheet.setCell("B3", ExcelCellValue.text("Done"));
      controller.setTable(
          workbook,
          new ExcelTableDefinition(
              "OpsTable",
              "V",
              "A1:B3",
              true,
              new ExcelTableStyle.Named("TableStyleMedium2", true, true, true, true)));

      XSSFTable table = workbook.sheet("V").xssfSheet().getTables().getFirst();
      var headerCell = workbook.sheet("V").xssfSheet().getRow(0).getCell(0);
      var staleStyle = workbook.xssfWorkbook().createCellStyle();
      staleStyle.cloneStyleFrom(headerCell.getCellStyle());
      staleStyle.setDataFormat(workbook.xssfWorkbook().createDataFormat().getFormat("yyyy-mm-dd"));
      headerCell.setCellStyle(staleStyle);

      assertEquals(
          "2026-02-06 00:19:17",
          table.getCTTable().getTableColumns().getTableColumnArray(0).getName());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertEquals(
          List.of("2026-02-06", "2026-06-26"),
          controller.tables(reopened, new ExcelTableSelection.All()).getFirst().columnNames());
    }
  }

  private static void populateTableCells(
      ExcelSheet sheet, String firstHeader, String secondHeader) {
    sheet.setCell("A1", ExcelCellValue.text(firstHeader));
    sheet.setCell("B1", ExcelCellValue.text(secondHeader));
    sheet.setCell("A2", ExcelCellValue.text("Ada"));
    sheet.setCell("B2", ExcelCellValue.text("Queue"));
    sheet.setCell("A3", ExcelCellValue.text("Lin"));
    sheet.setCell("B3", ExcelCellValue.text("Pack"));
  }
}
