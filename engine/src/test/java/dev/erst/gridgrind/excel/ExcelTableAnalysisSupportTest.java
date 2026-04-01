package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.junit.jupiter.api.Test;

/** Tests for package-private table analysis helpers. */
class ExcelTableAnalysisSupportTest {
  @Test
  void tableFindingsCoverBrokenHealthyAndStyleMismatchStates() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      List<WorkbookAnalysis.AnalysisFindingCode> brokenCodes =
          ExcelTableAnalysisSupport.tableFindings(
                  workbook,
                  new ExcelTableSnapshot(
                      "Queue",
                      "Ops",
                      "A1:B2",
                      0,
                      1,
                      List.of("", "owner", "Owner"),
                      new ExcelTableStyleSnapshot.Named("", false, false, true, false),
                      true))
              .stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();

      assertTrue(brokenCodes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_BROKEN_REFERENCE));
      assertTrue(brokenCodes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_BLANK_HEADER));
      assertTrue(brokenCodes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_DUPLICATE_HEADER));
      assertTrue(brokenCodes.contains(WorkbookAnalysis.AnalysisFindingCode.TABLE_STYLE_MISMATCH));

      assertEquals(
          List.of(),
          ExcelTableAnalysisSupport.tableFindings(
              workbook,
              new ExcelTableSnapshot(
                  "Queue",
                  "Ops",
                  "A1:B3",
                  1,
                  0,
                  List.of("Owner", "Task"),
                  new ExcelTableStyleSnapshot.None(),
                  true)));
    }
  }

  @Test
  void overlapFindingsIgnoreInvalidAndDifferentSheetTablesButReportRealOverlaps() {
    ExcelTableSnapshot subject =
        new ExcelTableSnapshot(
            "Queue",
            "Ops",
            "A1:B3",
            1,
            0,
            List.of("Owner", "Task"),
            new ExcelTableStyleSnapshot.None(),
            true);

    List<WorkbookAnalysis.AnalysisFinding> findings =
        ExcelTableAnalysisSupport.overlapFindings(
            subject,
            List.of(
                subject,
                new ExcelTableSnapshot(
                    "Invalid",
                    "Ops",
                    "A0:B3",
                    1,
                    0,
                    List.of("Owner", "Task"),
                    new ExcelTableStyleSnapshot.None(),
                    false),
                new ExcelTableSnapshot(
                    "ArchiveQueue",
                    "Archive",
                    "A1:B3",
                    1,
                    0,
                    List.of("Owner", "Task"),
                    new ExcelTableStyleSnapshot.None(),
                    false),
                new ExcelTableSnapshot(
                    "Desk",
                    "Ops",
                    "B1:C3",
                    1,
                    0,
                    List.of("Desk", "Region"),
                    new ExcelTableStyleSnapshot.None(),
                    false),
                new ExcelTableSnapshot(
                    "North",
                    "Ops",
                    "D1:E3",
                    1,
                    0,
                    List.of("Desk", "Region"),
                    new ExcelTableStyleSnapshot.None(),
                    false)));

    assertEquals(1, findings.size());
    assertEquals(
        WorkbookAnalysis.AnalysisFindingCode.TABLE_OVERLAPPING_RANGE, findings.getFirst().code());
  }

  @Test
  void tableAutofilterFindingsCoverHealthyInvalidMismatchAndBlankHeaderStates() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setCell("A1", ExcelCellValue.text("Owner"));
      sheet.setCell("B1", ExcelCellValue.text("Task"));
      sheet.setCell("A2", ExcelCellValue.text("Ada"));
      sheet.setCell("B2", ExcelCellValue.text("Queue"));
      sheet.setCell("A3", ExcelCellValue.text("Lin"));
      sheet.setCell("B3", ExcelCellValue.text("Pack"));
      sheet.setCell("A4", ExcelCellValue.text("Totals"));
      sheet.setCell("B4", ExcelCellValue.text("Done"));

      ExcelTableController controller = new ExcelTableController();
      controller.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B4", true, new ExcelTableStyle.None()));

      XSSFSheet poiSheet = workbook.sheet("Ops").xssfSheet();
      XSSFTable healthyTable = poiSheet.getTables().getFirst();
      ExcelTableSnapshot healthySnapshot =
          controller.tables(workbook, new ExcelTableSelection.All()).getFirst();
      List<WorkbookAnalysis.AnalysisFinding> healthyFindings =
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot);
      assertEquals(List.of(), healthyFindings, healthyFindings.toString());

      healthyTable.getCTTable().getAutoFilter().setRef("A0:B3");
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot)
              .getFirst()
              .code());

      healthyTable.getCTTable().getAutoFilter().setRef("A1:B4");
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot)
              .getFirst()
              .code());

      healthyTable.getCTTable().getAutoFilter().setRef("A1:B3");
      poiSheet.getRow(0).getCell(0).setBlank();
      poiSheet.getRow(0).getCell(1).setBlank();
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot)
              .getFirst()
              .code());
    }
  }
}
