package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.customxml.*;
import dev.erst.gridgrind.core.semantics.ExecutionConventions;
import dev.erst.gridgrind.engine.customxml.MappedDateInputException;
import dev.erst.gridgrind.engine.runtime.NativeInvocationContext;
import java.util.Optional;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Every mapped temporal row is admitted before any ordinary field or table size mutates. */
class MappedDateAdmissionTest {
  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  void lateInvalidTableDateRefusesBeforeAnyCellOrTableMutation(boolean mac) throws Exception {
    try (var _ = NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD);
        var workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      workbook.xssfWorkbook().getCTWorkbook().getWorkbookPr().setDate1904(mac);
      var table = table(workbook);
      String original = workbook.xssfWorkbook().getSheetAt(0).getCTWorksheet().xmlText();
      String rows = table.getXSSFSheet().getCTWorksheet().xmlText();
      String tableXml = table.getCTTable().xmlText();
      var failure =
          assertThrows(
              MappedDateInputException.class,
              () ->
                  workbook
                      .customXml()
                      .importMapping(
                          input(
                              "<CORSO><NOME>otherwise"
                                  + " changed</NOME><ROW><NAME>first</NAME><DATE>2026-03-08</DATE></ROW><ROW><NAME>second</NAME><DATE>1880-01-01</DATE></ROW></CORSO>")));
      assertEquals("/CORSO/ROW/DATE", failure.xpath());
      assertEquals(1, failure.mappingRowIndex());
      assertEquals(original, workbook.xssfWorkbook().getSheetAt(0).getCTWorksheet().xmlText());
      assertEquals(rows, table.getXSSFSheet().getCTWorksheet().xmlText());
      assertEquals(tableXml, table.getCTTable().xmlText());
      workbook
          .customXml()
          .importMapping(
              input(
                  "<CORSO><ROW><NAME>first</NAME><DATE>2026-03-08</DATE></ROW><ROW><NAME>second</NAME><DATE>2026-03-09</DATE></ROW></CORSO>"));
      assertEquals(
          mac ? 44627 : 46089, table.getXSSFSheet().getRow(1).getCell(1).getNumericCellValue());
      assertEquals(
          mac ? 44628 : 46090, table.getXSSFSheet().getRow(2).getCell(1).getNumericCellValue());
    }
  }

  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  void removedCurrentTableBindingIsANoOpInsteadOfAStaleTemporalRefusal(boolean mac)
      throws Exception {
    try (var _ = NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD);
        var workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      workbook.xssfWorkbook().getCTWorkbook().getWorkbookPr().setDate1904(mac);
      var table = table(workbook);
      table.getXSSFSheet().removeTable(table);
      String original = table.getXSSFSheet().getCTWorksheet().xmlText();
      workbook
          .customXml()
          .importMapping(
              input("<CORSO><ROW><NAME>unused</NAME><DATE>1880-01-01</DATE></ROW></CORSO>"));
      assertEquals(original, table.getXSSFSheet().getCTWorksheet().xmlText());
    }
  }

  private static ResolvedCustomXmlImport input(String xml) {
    return new ResolvedCustomXmlImport(
        new CustomXmlMappingLocator(Optional.of(1L), Optional.of("CORSO_mapping")), xml);
  }

  private static XSSFTable table(XssfWorkbook workbook) {
    var sheet = workbook.xssfWorkbook().createSheet("Rows");
    var header = sheet.createRow(0);
    header.createCell(0).setCellValue("Name");
    header.createCell(1).setCellValue("Date");
    var initial = sheet.createRow(1);
    initial.createCell(0).setCellValue("old");
    initial.createCell(1).setCellValue("old");
    var table = sheet.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
    table.setName("Dates");
    table.setDisplayName("Dates");
    for (int i = 0; i < 2; i++) {
      var binding = table.getCTTable().getTableColumns().getTableColumnArray(i).addNewXmlColumnPr();
      binding.setMapId(1);
      binding.setXpath(i == 0 ? "/CORSO/ROW/NAME" : "/CORSO/ROW/DATE");
      binding.setXmlDataType(i == 0 ? "string" : "date");
    }
    assertEquals("/CORSO/ROW", table.getCommonXpath());
    return table;
  }
}
