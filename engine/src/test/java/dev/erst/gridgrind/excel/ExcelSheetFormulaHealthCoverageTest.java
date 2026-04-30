package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** ExcelSheet formula readback and formula-health coverage. */
class ExcelSheetFormulaHealthCoverageTest extends ExcelSheetTestSupport {
  @Test
  void wrapsFormulaFailuresDuringReadsAndSnapshots() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      Row row = poiSheet.createRow(0);
      row.createCell(0).setCellFormula("1+1");

      ExcelSheet writeFailureSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              poiWorkbook.getCreationHelper().createFormulaEvaluator());
      InvalidFormulaException invalidWrite =
          assertThrows(
              InvalidFormulaException.class,
              () -> writeFailureSheet.setCell("B1", ExcelCellValue.formula("SUM(")));
      assertEquals("SUM(", invalidWrite.formula());
      InvalidFormulaException parserStateWrite =
          assertThrows(
              InvalidFormulaException.class,
              () -> writeFailureSheet.setCell("C1", ExcelCellValue.formula("[^owe_e`ffffff")));
      assertEquals("[^owe_e`ffffff", parserStateWrite.formula());

      FormulaEvaluator baseEvaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet invalidFormulaSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.failingEvaluation(
                  baseEvaluator, new org.apache.poi.ss.formula.FakeFormulaFailure("bad formula")));

      InvalidFormulaException invalidSnapshot =
          assertThrows(InvalidFormulaException.class, () -> invalidFormulaSheet.snapshotCell("A1"));
      assertEquals("1+1", invalidSnapshot.formula());
      InvalidFormulaException invalidNumber =
          assertThrows(InvalidFormulaException.class, () -> invalidFormulaSheet.number("A1"));
      assertEquals("1+1", invalidNumber.formula());
      InvalidFormulaException invalidBoolean =
          assertThrows(InvalidFormulaException.class, () -> invalidFormulaSheet.bool("A1"));
      assertEquals("1+1", invalidBoolean.formula());

      ExcelSheet displayFailureSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.alwaysFail(
                  new org.apache.poi.ss.formula.FakeFormulaFailure("display failure")));
      InvalidFormulaException displayFailure =
          assertThrows(InvalidFormulaException.class, () -> displayFailureSheet.snapshotCell("A1"));
      assertEquals("1+1", displayFailure.formula());

      ExcelSheet unsupportedFormulaSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.alwaysFail(
                  new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException(
                      "unsupported")));
      UnsupportedFormulaException unsupported =
          assertThrows(
              UnsupportedFormulaException.class, () -> unsupportedFormulaSheet.snapshotCell("A1"));
      assertEquals("1+1", unsupported.formula());

      ExcelSheet nullEvaluatedCellSheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.nullEvaluation(baseEvaluator));
      ExcelCellSnapshot blankEvaluatedFormula = nullEvaluatedCellSheet.snapshotCell("A1");
      assertEquals("FORMULA", blankEvaluatedFormula.effectiveType());
      assertThrows(IllegalStateException.class, () -> nullEvaluatedCellSheet.number("A1"));
      assertThrows(IllegalStateException.class, () -> nullEvaluatedCellSheet.bool("A1"));
      assertInstanceOf(
          ExcelCellSnapshot.BlankSnapshot.class,
          ((ExcelCellSnapshot.FormulaSnapshot) blankEvaluatedFormula).evaluation());
    }
  }

  @Test
  void previewRowReturnsEmptyCellsForGapRows() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Sparse");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.text("first"));
      sheet.setCell("A3", ExcelCellValue.text("third"));

      List<ExcelPreviewRow> preview = sheet.preview(3, 1);
      assertEquals(3, preview.size());
      assertEquals(1, preview.get(0).cells().size());
      assertEquals(0, preview.get(1).cells().size());
      assertEquals(1, preview.get(2).cells().size());
    }
  }

  @Test
  void derivesFormulaHealthFindingsAcrossExternalVolatileErrorAndFailureCases() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      FormulaEvaluator evaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet sheet =
          new ExcelSheet(poiSheet, new WorkbookStyleRegistry(poiWorkbook), evaluator);

      sheet.setCell("A1", ExcelCellValue.formula("INDIRECT(\"[External.xlsx]Sheet1!A1\")"));
      sheet.setCell("A2", ExcelCellValue.formula("NOW()"));
      sheet.setCell("A3", ExcelCellValue.formula("1/0"));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(3, sheet.formulaCellCount());
      assertTrue(
          findings.stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK,
                      AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                      AnalysisFindingCode.FORMULA_ERROR_RESULT)));
    }
  }

  @Test
  void derivesFormulaEvaluationFailureFindingWithFallbackExceptionMessage() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiSheet.createRow(0).createCell(0).setCellFormula("1+1");
      ExcelSheet sheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.alwaysFail(new IllegalStateException()));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(1, findings.size());
      WorkbookAnalysis.AnalysisFinding finding = findings.getFirst();
      assertEquals(AnalysisFindingCode.FORMULA_EVALUATION_FAILURE, finding.code());
      assertEquals(List.of("1+1", "IllegalStateException"), finding.evidence());
      assertEquals("Formula evaluation failed: IllegalStateException", finding.message());
    }
  }

  @Test
  void formulaHealthFindingsIgnoreSuccessfulAndNullEvaluations() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet successSheet = poiWorkbook.createSheet("Success");
      successSheet.createRow(0).createCell(0).setCellFormula("1+1");
      ExcelSheet successful =
          new ExcelSheet(
              successSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              poiWorkbook.getCreationHelper().createFormulaEvaluator());
      assertEquals(List.of(), successful.formulaHealthFindings());

      Sheet nullSheet = poiWorkbook.createSheet("NullEval");
      nullSheet.createRow(0).createCell(0).setCellFormula("1+1");
      FormulaEvaluator baseEvaluator = poiWorkbook.getCreationHelper().createFormulaEvaluator();
      ExcelSheet nullEvaluated =
          new ExcelSheet(
              nullSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.nullEvaluation(baseEvaluator));
      assertEquals(List.of(), nullEvaluated.formulaHealthFindings());
    }
  }

  @Test
  void formulaHealthFindingsDoNotDuplicateStaticMissingExternalWorkbookFindings() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      poiWorkbook.setCellFormulaValidation(false);
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiSheet.createRow(0).createCell(0).setCellFormula("[Rates.xlsx]Sheet1!A1");
      ExcelSheet sheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.alwaysFail(
                  new IllegalStateException(
                      "outer",
                      new LocalWorkbookNotFoundException(
                          "Could not resolve external workbook name 'Rates.xlsx'."))));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(1, findings.size());
      assertEquals(
          AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK, findings.getFirst().code());
    }
  }

  @Test
  void formulaHealthFindingsDoNotDuplicateStaticUnregisteredUdfFindings() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiSheet.createRow(0).createCell(0).setCellValue(21.0d);
      poiSheet.getRow(0).createCell(1).setCellFormula("DOUBLE(A1)");
      ExcelSheet sheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              FormulaRuntimeTestDouble.alwaysFail(
                  new org.apache.poi.ss.formula.eval.FakeNotImplementedFunctionException(
                      "DOUBLE")));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(1, findings.size());
      assertEquals(
          AnalysisFindingCode.FORMULA_UNREGISTERED_USER_DEFINED_FUNCTION,
          findings.getFirst().code());
    }
  }

  @Test
  void formulaHealthFindingsRespectBoundAndCachedExternalWorkbookPolicies() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      poiWorkbook.setCellFormulaValidation(false);

      Sheet boundSheet = poiWorkbook.createSheet("Bound");
      boundSheet.createRow(0).createCell(0).setCellFormula("[Rates.xlsx]Sheet1!A1");
      ExcelSheet bound =
          new ExcelSheet(
              boundSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              new StaticContextFormulaRuntime(
                  new ExcelFormulaRuntimeContext(
                      Set.of("Rates.xlsx"), ExcelFormulaMissingWorkbookPolicy.ERROR, Set.of()),
                  null,
                  null));
      assertEquals(List.of(), bound.formulaHealthFindings());

      Sheet cachedSheet = poiWorkbook.createSheet("Cached");
      cachedSheet.createRow(0).createCell(0).setCellFormula("[Forecast.xlsx]Sheet1!A1");
      ExcelSheet cached =
          new ExcelSheet(
              cachedSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              new StaticContextFormulaRuntime(
                  new ExcelFormulaRuntimeContext(
                      Set.of(), ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE, Set.of()),
                  null,
                  null));

      List<WorkbookAnalysis.AnalysisFinding> findings = cached.formulaHealthFindings();

      assertEquals(1, findings.size());
      assertEquals(
          AnalysisFindingCode.FORMULA_USES_CACHED_EXTERNAL_VALUE, findings.getFirst().code());
      assertEquals(AnalysisSeverity.INFO, findings.getFirst().severity());
      assertEquals(
          List.of("[Forecast.xlsx]Sheet1!A1", "Forecast.xlsx"), findings.getFirst().evidence());
    }
  }

  @Test
  void formulaHealthFindingsReportRuntimeMissingExternalWorkbookWhenStaticScanCannotInferIt()
      throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiSheet.createRow(0).createCell(0).setCellFormula("SUM(A2)");
      ExcelSheet sheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              new StaticContextFormulaRuntime(
                  ExcelFormulaEnvironment.defaults().runtimeContext(),
                  new IllegalStateException(
                      "outer",
                      new LocalWorkbookNotFoundException(
                          "Could not resolve external workbook name 'Rates.xlsx'.")),
                  null));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(1, findings.size());
      assertEquals(
          AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK, findings.getFirst().code());
      assertEquals(List.of("SUM(A2)", "Rates.xlsx"), findings.getFirst().evidence());
    }
  }

  @Test
  void formulaHealthFindingsLeaveMissingWorkbookEvidenceBlankWhenWorkbookNameCannotBeDerived()
      throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook()) {
      Sheet poiSheet = poiWorkbook.createSheet("Budget");
      poiSheet.createRow(0).createCell(0).setCellFormula("SUM(A2)");
      ExcelSheet sheet =
          new ExcelSheet(
              poiSheet,
              new WorkbookStyleRegistry(poiWorkbook),
              new StaticContextFormulaRuntime(
                  ExcelFormulaEnvironment.defaults().runtimeContext(),
                  new IllegalStateException("outer", new LocalWorkbookNotFoundException(null)),
                  null));

      List<WorkbookAnalysis.AnalysisFinding> findings = sheet.formulaHealthFindings();

      assertEquals(1, findings.size());
      assertEquals(
          AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK, findings.getFirst().code());
      assertEquals(List.of("SUM(A2)", ""), findings.getFirst().evidence());
    }
  }
}
