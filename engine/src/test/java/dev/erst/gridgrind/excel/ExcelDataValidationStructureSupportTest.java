package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

/** Tests for data-validation normalization during supported structural inserts. */
class ExcelDataValidationStructureSupportTest {
  @Test
  void rowInsertDropsMalformedValidationsWithBlankSqref() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(STDataValidationType.LIST, List.of(" "), "\"Queued\""));

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(sheet, 0, 1);

      assertTrue(rewritten.isEmpty());
    }
  }

  @Test
  void columnInsertDropsMalformedValidationsWithBlankSqref() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(STDataValidationType.LIST, List.of(" "), "\"Queued\""));

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertColumns(sheet, 0, 1);

      assertTrue(rewritten.isEmpty());
    }
  }

  @Test
  void rowInsertRetainsUntypedAndBlankFormulaValidations() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(null, List.of("A1"), null));

      CTDataValidation blankFormulaValidation =
          validation(STDataValidationType.WHOLE, List.of("B1"), "");
      blankFormulaValidation.setFormula2("");
      blankFormulaValidation.setOperator(STDataValidationOperator.BETWEEN);
      addRawValidation(sheet, blankFormulaValidation);

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(sheet, 0, 1);

      assertEquals(2, rewritten.size());
      assertEquals(List.of("A2"), normalizedSqref(rewritten.getFirst()));
      assertFalse(rewritten.getFirst().isSetType());
      assertFalse(rewritten.getFirst().isSetFormula1());
      assertEquals(List.of("B2"), normalizedSqref(rewritten.get(1)));
      assertEquals("", rewritten.get(1).getFormula1());
      assertEquals("", rewritten.get(1).getFormula2());
    }
  }

  @Test
  void rowInsertRetainsListValidationsWithoutFormulaAndQuotedLiterals() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(STDataValidationType.LIST, List.of("A1:A2"), null));
      addRawValidation(sheet, validation(STDataValidationType.LIST, List.of("B1:B2"), "\"0,1\""));

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(sheet, 0, 1);

      assertEquals(2, rewritten.size());
      assertEquals(List.of("A2:A3"), normalizedSqref(rewritten.getFirst()));
      assertFalse(rewritten.getFirst().isSetFormula1());
      assertEquals(List.of("B2:B3"), normalizedSqref(rewritten.get(1)));
      assertEquals("\"0,1\"", rewritten.get(1).getFormula1());
    }
  }

  @Test
  void rowInsertLeavesNamedListFormulasUnchangedWhenNoReferencesMove() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet lookup = workbook.createSheet("Lookup");
      lookup.createRow(0).createCell(0).setCellValue("Queued");
      lookup.getRow(0).createCell(1).setCellValue("Done");
      Name statuses = workbook.createName();
      statuses.setNameName("Statuses");
      statuses.setRefersToFormula("Lookup!$A$1:$B$1");

      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(STDataValidationType.LIST, List.of("A1:A2"), "Statuses"));

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(sheet, 0, 1);

      assertEquals(1, rewritten.size());
      assertEquals(List.of("A2:A3"), normalizedSqref(rewritten.getFirst()));
      assertEquals("Statuses", rewritten.getFirst().getFormula1());
    }
  }

  @Test
  void rowInsertTreatsSingleCharacterListFormulasAsNonLiteralExpressions() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet lookup = workbook.createSheet("Lookup");
      lookup.createRow(0).createCell(0).setCellValue("Queued");
      Name q = workbook.createName();
      q.setNameName("Q");
      q.setRefersToFormula("Lookup!$A$1");

      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(STDataValidationType.LIST, List.of("A1:A2"), "Q"));

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(sheet, 0, 1);

      assertEquals(1, rewritten.size());
      assertEquals("Q", rewritten.getFirst().getFormula1());
    }
  }

  @Test
  void rowInsertRejectsListFormulasWithAnOpeningQuoteOnly() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(STDataValidationType.LIST, List.of("A1:A2"), "\"Queued"));

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(
                      sheet, 0, 1));

      assertTrue(
          failure.getMessage().contains("Cannot normalize data-validation formula '\"Queued'"));
    }
  }

  @Test
  void rowInsertShiftsFormulaBoundsForCellBasedRules() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      CTDataValidation validation = validation(STDataValidationType.WHOLE, List.of("C1"), "A1");
      validation.setFormula2("B1");
      validation.setOperator(STDataValidationOperator.BETWEEN);
      addRawValidation(sheet, validation);

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(sheet, 0, 1);

      assertEquals(1, rewritten.size());
      assertEquals("A2", rewritten.getFirst().getFormula1());
      assertEquals("B2", rewritten.getFirst().getFormula2());
    }
  }

  @Test
  void columnInsertSplitsIntersectingValidationRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(
          sheet, validation(STDataValidationType.LIST, List.of("A1:C3"), "\"Queued\""));

      List<CTDataValidation> rewritten =
          ExcelDataValidationStructureSupport.expectedValidationsAfterInsertColumns(sheet, 1, 1);

      assertEquals(1, rewritten.size());
      assertEquals(List.of("A1:A3", "C1:D3"), normalizedSqref(rewritten.getFirst()));
      assertEquals("\"Queued\"", rewritten.getFirst().getFormula1());
    }
  }

  @Test
  void rowInsertRejectsValidationOverflowAtWorksheetEdge() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(
          sheet,
          validation(
              STDataValidationType.LIST,
              List.of("A" + (ExcelRowSpan.MAX_ROW_INDEX + 1)),
              "\"Queued\""));

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(
                      sheet, ExcelRowSpan.MAX_ROW_INDEX, 1));

      assertTrue(failure.getMessage().contains("beyond the maximum row"));
    }
  }

  @Test
  void columnInsertRejectsValidationOverflowAtWorksheetEdge() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(
          sheet,
          validation(
              STDataValidationType.LIST,
              List.of(SpreadsheetVersion.EXCEL2007.getLastColumnName() + "1"),
              "\"Queued\""));

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelDataValidationStructureSupport.expectedValidationsAfterInsertColumns(
                      sheet, ExcelColumnSpan.MAX_COLUMN_INDEX, 1));

      assertTrue(failure.getMessage().contains("beyond the maximum column"));
    }
  }

  @Test
  void rowInsertRejectsMalformedValidationFormulasWithContext() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, validation(STDataValidationType.CUSTOM, List.of("A1"), "SUM("));

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelDataValidationStructureSupport.expectedValidationsAfterInsertRows(
                      sheet, 0, 1));

      assertTrue(failure.getMessage().contains("Cannot normalize data-validation formula 'SUM('"));
      assertTrue(failure.getMessage().contains("sheet 'Budget'"));
    }
  }

  private static List<String> normalizedSqref(CTDataValidation validation) {
    return ExcelSqrefSupport.normalizedSqref(validation.getSqref());
  }

  private static void addRawValidation(XSSFSheet sheet, CTDataValidation validation) {
    CTDataValidations dataValidations =
        sheet.getCTWorksheet().isSetDataValidations()
            ? sheet.getCTWorksheet().getDataValidations()
            : sheet.getCTWorksheet().addNewDataValidations();
    dataValidations.addNewDataValidation().set(validation);
    dataValidations.setCount(dataValidations.sizeOfDataValidationArray());
  }

  private static CTDataValidation validation(
      STDataValidationType.Enum type, List<String> sqref, String formula1) {
    CTDataValidation validation = CTDataValidation.Factory.newInstance();
    if (type != null) {
      validation.setType(type);
    }
    validation.setSqref(sqref);
    if (formula1 != null) {
      validation.setFormula1(formula1);
    }
    return validation;
  }
}
