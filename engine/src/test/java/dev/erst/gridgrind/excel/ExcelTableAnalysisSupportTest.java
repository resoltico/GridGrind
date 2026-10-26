package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.junit.jupiter.api.Test;

/** Tests for package-private table analysis helpers. */
class ExcelTableAnalysisSupportTest {
  @Test
  void tableFindingsCoverBrokenHealthyAndStyleMismatchStates() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      List<AnalysisFindingCode> brokenCodes =
          ExcelTableAnalysisSupport.tableFindings(
                  workbook,
                  new TableEntryReport(
                      new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                      new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"),
                      0,
                      1,
                      List.of("", "owner", "Owner"),
                      new TableStyleReport.Named("", false, false, true, false),
                      true))
              .stream()
              .map(AnalysisFindingReport::code)
              .toList();

      assertTrue(brokenCodes.contains(AnalysisFindingCode.TABLE_BROKEN_REFERENCE));
      assertTrue(brokenCodes.contains(AnalysisFindingCode.TABLE_BLANK_HEADER));
      assertTrue(brokenCodes.contains(AnalysisFindingCode.TABLE_DUPLICATE_HEADER));
      assertTrue(brokenCodes.contains(AnalysisFindingCode.TABLE_STYLE_MISMATCH));

      assertEquals(
          List.of(),
          ExcelTableAnalysisSupport.tableFindings(
              workbook,
              new TableEntryReport(
                  new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                  new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                  1,
                  0,
                  List.of("Owner", "Task"),
                  new TableStyleReport.None(),
                  true)));
    }
  }

  @Test
  void overlapFindingsIgnoreInvalidAndDifferentSheetTablesButReportRealOverlaps() {
    TableEntryReport subject =
        new TableEntryReport(
            new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
            1,
            0,
            List.of("Owner", "Task"),
            new TableStyleReport.None(),
            true);

    List<AnalysisFindingReport> findings =
        ExcelTableAnalysisSupport.overlapFindings(
            subject,
            List.of(
                subject,
                new TableEntryReport(
                    new dev.erst.gridgrind.core.vocabulary.TableLabel("Invalid"),
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A0:B3"),
                    1,
                    0,
                    List.of("Owner", "Task"),
                    new TableStyleReport.None(),
                    false),
                new TableEntryReport(
                    new dev.erst.gridgrind.core.vocabulary.TableLabel("ArchiveQueue"),
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Archive"),
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                    1,
                    0,
                    List.of("Owner", "Task"),
                    new TableStyleReport.None(),
                    false),
                new TableEntryReport(
                    new dev.erst.gridgrind.core.vocabulary.TableLabel("Desk"),
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("B1:C3"),
                    1,
                    0,
                    List.of("Desk", "Region"),
                    new TableStyleReport.None(),
                    false),
                new TableEntryReport(
                    new dev.erst.gridgrind.core.vocabulary.TableLabel("North"),
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("D1:E3"),
                    1,
                    0,
                    List.of("Desk", "Region"),
                    new TableStyleReport.None(),
                    false)));

    assertEquals(1, findings.size());
    assertEquals(AnalysisFindingCode.TABLE_OVERLAPPING_RANGE, findings.getFirst().code());
  }

  @Test
  void tableAutofilterFindingsCoverHealthyInvalidMismatchAndBlankHeaderStates() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.cells().setCell("A1", new CellValue.Text("Owner"));
      sheet.cells().setCell("B1", new CellValue.Text("Task"));
      sheet.cells().setCell("A2", new CellValue.Text("Ada"));
      sheet.cells().setCell("B2", new CellValue.Text("Queue"));
      sheet.cells().setCell("A3", new CellValue.Text("Lin"));
      sheet.cells().setCell("B3", new CellValue.Text("Pack"));
      sheet.cells().setCell("A4", new CellValue.Text("Totals"));
      sheet.cells().setCell("B4", new CellValue.Text("Done"));

      ExcelTableController controller = new ExcelTableController();
      controller.setTable(
          workbook,
          new ResolvedTable(
              new TableName("Queue"),
              new SheetName("Ops"),
              new RangeRef("A1:B4"),
              true,
              new TableStyleSpec.None()));

      XSSFSheet poiSheet = workbook.sheet("Ops").xssfSheet();
      XSSFTable healthyTable = poiSheet.getTables().getFirst();
      TableEntryReport healthySnapshot =
          controller.tables(workbook, new TableSelector.All()).getFirst();
      List<AnalysisFindingReport> healthyFindings =
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot);
      assertEquals(List.of(), healthyFindings, healthyFindings.toString());

      healthyTable.getCTTable().getAutoFilter().setRef("A0:B3");
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot)
              .getFirst()
              .code());

      healthyTable.getCTTable().getAutoFilter().setRef("A1:B4");
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot)
              .getFirst()
              .code());

      healthyTable.getCTTable().getAutoFilter().setRef("A1:B3");
      poiSheet.getRow(0).getCell(0).setBlank();
      poiSheet.getRow(0).getCell(1).setBlank();
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
          ExcelTableAnalysisSupport.tableAutofilterFindings(poiSheet, healthySnapshot)
              .getFirst()
              .code());
    }
  }
}
