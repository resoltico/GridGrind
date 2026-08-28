package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Verifies Unicode defined names and table names survive authored and existing-workbook reads. */
class ExcelUnicodeDefinedNamePersistenceTest {
  private static final List<UnicodeNameFixture> FIXTURES =
      List.of(
          new UnicodeNameFixture("IeņēmumiRange", "IeņēmumiTable", "A1:B2"),
          new UnicodeNameFixture("ДоходыRange", "ДоходыTable", "D1:E2"),
          new UnicodeNameFixture("収益Range", "収益Table", "G1:H2"));

  @Test
  void authorsAndReopensUnicodeNamedRangesAndTables() throws IOException {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-unicode-author-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Data");
      populateTableRanges(sheet);
      for (UnicodeNameFixture fixture : FIXTURES) {
        authorFixture(workbook, fixture);
      }
      workbook
          .persistence()
          .save(
              workbookPath,
              WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(
          FIXTURES.stream().map(UnicodeNameFixture::namedRangeName).toList(),
          workbook.names().namedRanges().stream().map(ExcelNamedRangeSnapshot::name).toList());
      assertEquals(
          FIXTURES.stream().map(UnicodeNameFixture::tableName).toList(),
          workbook.xssfWorkbook().getSheet("Data").getTables().stream()
              .map(table -> table.getName())
              .toList());
    }
  }

  @Test
  void readsUnicodeNamesAndTablesCreatedByAnExistingWorkbook() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-unicode-observed-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      populateTableRanges(sheet);
      for (UnicodeNameFixture fixture : FIXTURES) {
        authorObservedFixture(workbook, sheet, fixture);
      }
      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertTrue(
          workbook.names().namedRanges().stream()
              .map(ExcelNamedRangeSnapshot::name)
              .collect(java.util.stream.Collectors.toSet())
              .containsAll(FIXTURES.stream().map(UnicodeNameFixture::namedRangeName).toList()));
      assertTrue(
          workbook.xssfWorkbook().getSheet("Data").getTables().stream()
              .map(table -> table.getName())
              .toList()
              .containsAll(FIXTURES.stream().map(UnicodeNameFixture::tableName).toList()));
    }
  }

  private static void populateTableRanges(ExcelSheet sheet) {
    for (UnicodeNameFixture fixture : FIXTURES) {
      populateTableRange(sheet, fixture.range());
    }
  }

  private static void populateTableRanges(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    for (UnicodeNameFixture fixture : FIXTURES) {
      populateObservedTableRange(sheet, fixture.range());
    }
  }

  private static String column(int columnIndex) {
    return org.apache.poi.ss.util.CellReference.convertNumToColString(columnIndex);
  }

  private static void authorFixture(ExcelWorkbook workbook, UnicodeNameFixture fixture) {
    workbook
        .names()
        .setNamedRange(
            new ExcelNamedRangeDefinition(
                fixture.namedRangeName(),
                new ExcelNamedRangeScope.WorkbookScope(),
                ExcelNamedRangeTarget.range("Data", fixture.range())));
    workbook
        .tables()
        .setTable(
            new ExcelTableDefinition(
                fixture.tableName(), "Data", fixture.range(), false, new ExcelTableStyle.None()));
  }

  private static void authorObservedFixture(
      XSSFWorkbook workbook,
      org.apache.poi.xssf.usermodel.XSSFSheet sheet,
      UnicodeNameFixture fixture) {
    ExcelRange parsed = ExcelRange.parse(fixture.range());
    var namedRange = workbook.createName();
    namedRange.setNameName(fixture.namedRangeName());
    namedRange.setRefersToFormula(
        "Data!$" + column(parsed.firstColumn()) + "$1:$" + column(parsed.lastColumn()) + "$2");
    var table = sheet.createTable(new AreaReference(fixture.range(), SpreadsheetVersion.EXCEL2007));
    table.setName(fixture.tableName());
    table.setDisplayName(fixture.tableName());
  }

  private static void populateTableRange(ExcelSheet sheet, String range) {
    ExcelRange parsed = ExcelRange.parse(range);
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.firstRow(), parsed.firstColumn())
                .formatAsString(),
            ExcelCellValue.text("Header" + parsed.firstColumn()));
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.firstRow(), parsed.lastColumn())
                .formatAsString(),
            ExcelCellValue.text("Value" + parsed.lastColumn()));
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.lastRow(), parsed.firstColumn())
                .formatAsString(),
            ExcelCellValue.text("one"));
    sheet
        .cells()
        .setCell(
            new org.apache.poi.ss.util.CellReference(parsed.lastRow(), parsed.lastColumn())
                .formatAsString(),
            ExcelCellValue.text("two"));
  }

  private static void populateObservedTableRange(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet, String range) {
    ExcelRange parsed = ExcelRange.parse(range);
    var header = sheet.getRow(parsed.firstRow());
    if (header == null) {
      header = sheet.createRow(parsed.firstRow());
    }
    header.createCell(parsed.firstColumn()).setCellValue("Header" + parsed.firstColumn());
    header.createCell(parsed.lastColumn()).setCellValue("Value" + parsed.lastColumn());
    var row = sheet.getRow(parsed.lastRow());
    if (row == null) {
      row = sheet.createRow(parsed.lastRow());
    }
    row.createCell(parsed.firstColumn()).setCellValue("one");
    row.createCell(parsed.lastColumn()).setCellValue("two");
  }

  private record UnicodeNameFixture(String namedRangeName, String tableName, String range) {}
}
