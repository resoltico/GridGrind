package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.MissingArgEval;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.RefEvalBase;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for formula-runtime helper coverage and failure handling. */
class ExcelFormulaRuntimeTest {
  @Test
  void formulaRuntimeRejectsInvalidTemplatesAndClosesOpenedBindingsOnSetupFailure()
      throws Exception {
    ExcelFormulaEnvironment invalidTemplateEnvironment =
        new ExcelFormulaEnvironment(
            List.of(),
            ExcelFormulaMissingWorkbookPolicy.ERROR,
            List.of(
                new ExcelFormulaUdfToolpack(
                    "math", List.of(new ExcelFormulaUdfFunction("BROKEN", 1, 1, "SUM(")))));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelFormulaRuntime.poi(workbook, invalidTemplateEnvironment));
      assertTrue(failure.getMessage().contains("Invalid formulaTemplate"));
    }

    Path directory = ExcelTempFiles.createManagedTempDirectory("gridgrind-formula-runtime-");
    Path validWorkbookPath = directory.resolve("valid.xlsx");
    Path missingWorkbookPath = directory.resolve("missing.xlsx");

    try (XSSFWorkbook validWorkbook = new XSSFWorkbook()) {
      validWorkbook.createSheet("Rates").createRow(0).createCell(0).setCellValue(7.5d);
      try (var outputStream = Files.newOutputStream(validWorkbookPath)) {
        validWorkbook.write(outputStream);
      }
    }

    ExcelFormulaEnvironment failingBindingEnvironment =
        new ExcelFormulaEnvironment(
            List.of(
                new ExcelFormulaExternalWorkbookBinding("valid.xlsx", validWorkbookPath),
                new ExcelFormulaExternalWorkbookBinding("missing.xlsx", missingWorkbookPath)),
            ExcelFormulaMissingWorkbookPolicy.ERROR,
            List.of());

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertThrows(
          IOException.class, () -> ExcelFormulaRuntime.poi(workbook, failingBindingEnvironment));
    }

    try (var reopened = WorkbookFactory.create(validWorkbookPath.toFile())) {
      assertEquals(7.5d, reopened.getSheet("Rates").getRow(0).getCell(0).getNumericCellValue());
    }
  }

  @Test
  void templateFormulaHelpersCoverSupportedAndRejectedValueEvalShapes() throws Exception {
    ExcelFormulaRuntime.TemplateFormulaUdfFunction identity =
        new ExcelFormulaRuntime.TemplateFormulaUdfFunction(
            new ExcelFormulaUdfFunction("IDENTITY", 1, 1, "ARG1"));
    ExcelFormulaRuntime.TemplateFormulaUdfFunction sumFunction =
        new ExcelFormulaRuntime.TemplateFormulaUdfFunction(
            new ExcelFormulaUdfFunction("SUM_AREA", 1, 1, "SUM(ARG1)"));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Args");
      Cell scalarCell = sheet.createRow(20).createCell(0);

      AreaEval areaEval = writeArgumentCoversScalarReferenceAndAreaShapes(sheet);
      writeScalarValueCoversSupportedAndRejectedShapes(scalarCell, areaEval);
      toValueEvalCoversSupportedAndRejectedShapes();
      evaluateCoversArgumentValidationAndAreaComputation(identity, sumFunction);
    }

    assertNotNull(identity);
  }

  @Test
  void referencedWorkbookHandleValidatesConstructorArguments() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      try (ExcelFormulaRuntime.ReferencedWorkbookHandle handle =
          new ExcelFormulaRuntime.ReferencedWorkbookHandle("Rates.xlsx", workbook, evaluator); ) {
        assertEquals("Rates.xlsx", handle.workbookName());
        assertSame(workbook, handle.workbook());
        assertSame(evaluator, handle.formulaEvaluator());

        assertThrows(
            NullPointerException.class,
            () -> new ExcelFormulaRuntime.ReferencedWorkbookHandle(null, workbook, evaluator));
        assertThrows(
            NullPointerException.class,
            () -> new ExcelFormulaRuntime.ReferencedWorkbookHandle("Rates.xlsx", null, evaluator));
        assertThrows(
            NullPointerException.class,
            () -> new ExcelFormulaRuntime.ReferencedWorkbookHandle("Rates.xlsx", workbook, null));
      }
    }
  }

  @Test
  void formulaRuntimeHelperMethodsReuseScratchStructuresAndGracefullyRejectInvalidInputs()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        ExcelFormulaRuntime runtime = ExcelFormulaRuntime.poi(workbook, null)) {
      assertEquals(ExcelFormulaEnvironment.defaults().runtimeContext(), runtime.context());
    }

    try (XSSFWorkbook scratchWorkbook = new XSSFWorkbook()) {
      XSSFSheet argsSheet = scratchWorkbook.createSheet("Args");
      argsSheet.createRow(0);
      argsSheet.createRow(1);

      ExcelFormulaRuntime.defineScratchArgument(scratchWorkbook, argsSheet, 0);
      ExcelFormulaRuntime.defineScratchArgument(scratchWorkbook, argsSheet, 1);

      assertEquals(1.0d, argsSheet.getRow(0).getCell(0).getNumericCellValue());
      assertEquals("Args!$A$1", scratchWorkbook.getName("_GRIDGRIND_ARG_1").getRefersToFormula());
      assertEquals(2.0d, argsSheet.getRow(1).getCell(0).getNumericCellValue());
      assertEquals("Args!$A$2", scratchWorkbook.getName("_GRIDGRIND_ARG_2").getRefersToFormula());

      Cell first = ExcelFormulaRuntime.TemplateFormulaUdfFunction.cell(argsSheet, 3, 2);
      Cell second = ExcelFormulaRuntime.TemplateFormulaUdfFunction.cell(argsSheet, 3, 2);
      assertSame(first, second);
    }

    ExcelFormulaRuntime.TemplateFormulaUdfFunction identity =
        new ExcelFormulaRuntime.TemplateFormulaUdfFunction(
            new ExcelFormulaUdfFunction("IDENTITY", 1, 1, "ARG1"));
    assertSame(ErrorEval.VALUE_INVALID, identity.evaluate(new ValueEval[] {}, null));
    assertSame(
        ErrorEval.VALUE_INVALID,
        identity.evaluate(new ValueEval[] {new NumberEval(1.0d), new NumberEval(2.0d)}, null));
    assertSame(
        ErrorEval.VALUE_INVALID,
        identity.evaluate(new ValueEval[] {new UnsupportedValueEval()}, null));
  }

  @SuppressWarnings({"PMD.CloseResource", "PMD.UseTryWithResources"})
  @Test
  void formulaRuntimeAggregatesReferencedWorkbookCloseFailures() throws Exception {
    ThrowingWorkbook firstReferencedWorkbook = new ThrowingWorkbook("first close failure");
    ThrowingWorkbook secondReferencedWorkbook = new ThrowingWorkbook("second close failure");
    XSSFWorkbook ownerWorkbook = new XSSFWorkbook();

    try {
      IOException referencedIOException =
          assertThrows(
              IOException.class,
              () ->
                  ExcelFormulaRuntime.closeReferencedWorkbooks(
                      List.of(
                          new ExcelFormulaRuntime.ReferencedWorkbookHandle(
                              "Rates.xlsx",
                              firstReferencedWorkbook,
                              firstReferencedWorkbook.getCreationHelper().createFormulaEvaluator()),
                          new ExcelFormulaRuntime.ReferencedWorkbookHandle(
                              "Taxes.xlsx",
                              secondReferencedWorkbook,
                              secondReferencedWorkbook
                                  .getCreationHelper()
                                  .createFormulaEvaluator()))));
      assertEquals("first close failure", referencedIOException.getMessage());
      assertEquals(1, referencedIOException.getSuppressed().length);
      assertEquals("second close failure", referencedIOException.getSuppressed()[0].getMessage());

      ExcelFormulaRuntime runtime =
          new ExcelFormulaRuntime.PoiExcelFormulaRuntime(
              ownerWorkbook.getCreationHelper().createFormulaEvaluator(),
              ExcelFormulaEnvironment.defaults().runtimeContext(),
              List.of(firstReferencedWorkbook, secondReferencedWorkbook));
      IOException runtimeCloseFailure = assertThrows(IOException.class, runtime::close);
      assertEquals("first close failure", runtimeCloseFailure.getMessage());
      assertEquals(1, runtimeCloseFailure.getSuppressed().length);
      assertEquals("second close failure", runtimeCloseFailure.getSuppressed()[0].getMessage());
    } finally {
      firstReferencedWorkbook.disableCloseFailure();
      secondReferencedWorkbook.disableCloseFailure();
      firstReferencedWorkbook.close();
      secondReferencedWorkbook.close();
      ownerWorkbook.close();
    }
  }

  private static AreaEval writeArgumentCoversScalarReferenceAndAreaShapes(XSSFSheet sheet) {
    var scalarAddress =
        ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeArgument(
            sheet, 0, new NumberEval(7.0d));
    assertEquals(0, scalarAddress.getFirstRow());
    assertEquals(7.0d, sheet.getRow(0).getCell(0).getNumericCellValue());

    var refAddress =
        ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeArgument(
            sheet, 2, new StubRefEval(2, 0, 0, new NumberEval(0.0d)));
    assertEquals(2, refAddress.getFirstRow());
    assertEquals(0.0d, sheet.getRow(2).getCell(0).getNumericCellValue());

    AreaEval areaEval =
        new StubAreaEval(
            4,
            5,
            0,
            1,
            new ValueEval[][] {
              {new NumberEval(1.0d), new StringEval("two")},
              {BoolEval.valueOf(true), ErrorEval.DIV_ZERO}
            });
    var areaAddress =
        ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeArgument(sheet, 4, areaEval);
    assertEquals(4, areaAddress.getFirstRow());
    assertEquals(5, areaAddress.getLastRow());
    assertEquals("two", sheet.getRow(4).getCell(1).getStringCellValue());
    assertTrue(sheet.getRow(5).getCell(0).getBooleanCellValue());
    assertEquals(FormulaError.DIV0.getCode(), sheet.getRow(5).getCell(1).getErrorCellValue());

    IllegalArgumentException multiSheetRefFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeArgument(
                    sheet, 8, new StubRefEval(8, 0, new NumberEval(0.0d), 0, 1)));
    assertTrue(multiSheetRefFailure.getMessage().contains("3D references"));
    return areaEval;
  }

  private static void writeScalarValueCoversSupportedAndRejectedShapes(
      Cell scalarCell, AreaEval areaEval) {
    assertDoesNotThrow(
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, BlankEval.instance));
    assertEquals(CellType.BLANK, scalarCell.getCellType());
    assertDoesNotThrow(
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, MissingArgEval.instance));
    assertEquals(CellType.BLANK, scalarCell.getCellType());
    assertDoesNotThrow(
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, new NumberEval(3.5d)));
    assertEquals(3.5d, scalarCell.getNumericCellValue());
    assertDoesNotThrow(
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, new StringEval("hello")));
    assertEquals("hello", scalarCell.getStringCellValue());
    assertDoesNotThrow(
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, BoolEval.valueOf(true)));
    assertTrue(scalarCell.getBooleanCellValue());
    assertDoesNotThrow(
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, ErrorEval.NAME_INVALID));
    assertEquals(FormulaError.NAME.getCode(), scalarCell.getErrorCellValue());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, new StubRefEval(0, 0, 0, new NumberEval(0.0d))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(scalarCell, areaEval));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(
                scalarCell, new UnsupportedValueEval()));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelFormulaRuntime.TemplateFormulaUdfFunction.writeScalarValue(scalarCell, null));
  }

  private static void toValueEvalCoversSupportedAndRejectedShapes() {
    assertSame(
        BlankEval.instance,
        ExcelFormulaRuntime.TemplateFormulaUdfFunction.toValueEval((CellValue) null));
    assertEquals(
        9.0d,
        ((NumberEval)
                ExcelFormulaRuntime.TemplateFormulaUdfFunction.toValueEval(new CellValue(9.0d)))
            .getNumberValue());
    assertEquals(
        "text",
        ((StringEval)
                ExcelFormulaRuntime.TemplateFormulaUdfFunction.toValueEval(new CellValue("text")))
            .getStringValue());
    assertTrue(
        ((BoolEval)
                ExcelFormulaRuntime.TemplateFormulaUdfFunction.toValueEval(CellValue.valueOf(true)))
            .getBooleanValue());
    assertEquals(
        FormulaError.REF.getCode(),
        ((ErrorEval)
                ExcelFormulaRuntime.TemplateFormulaUdfFunction.toValueEval(
                    CellValue.getError(FormulaError.REF.getCode())))
            .getErrorCode());
    assertSame(
        BlankEval.instance,
        ExcelFormulaRuntime.TemplateFormulaUdfFunction.toValueEval(
            CellType.BLANK, 0.0d, false, null, 0));
    IllegalStateException formulaCellValueFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelFormulaRuntime.TemplateFormulaUdfFunction.toValueEval(
                    CellType.FORMULA, 0.0d, false, null, 0));
    assertEquals("Scratch evaluator returned FORMULA", formulaCellValueFailure.getMessage());
  }

  private static void evaluateCoversArgumentValidationAndAreaComputation(
      ExcelFormulaRuntime.TemplateFormulaUdfFunction identity,
      ExcelFormulaRuntime.TemplateFormulaUdfFunction sumFunction) {
    assertSame(ErrorEval.VALUE_INVALID, identity.evaluate(new ValueEval[0], null));
    ValueEval identityResult = identity.evaluate(new ValueEval[] {new NumberEval(6.0d)}, null);
    assertInstanceOf(NumberEval.class, identityResult);
    assertEquals(6.0d, ((NumberEval) identityResult).getNumberValue());
    ValueEval sumResult =
        sumFunction.evaluate(
            new ValueEval[] {
              new StubAreaEval(
                  0, 1, 0, 0, new ValueEval[][] {{new NumberEval(2.0d)}, {new NumberEval(4.0d)}})
            },
            null);
    assertInstanceOf(NumberEval.class, sumResult);
    assertEquals(6.0d, ((NumberEval) sumResult).getNumberValue());
  }

  /** Test-only unsupported value shape used to hit invalid-evaluator branches. */
  private static final class UnsupportedValueEval implements ValueEval {}

  /** Test-only `RefEval` stub with controllable scalar and 3D reference behavior. */
  private static final class StubRefEval extends RefEvalBase {
    private final ValueEval valueEval;

    private StubRefEval(int rowIndex, int columnIndex, int sheetIndex, ValueEval valueEval) {
      this(rowIndex, columnIndex, valueEval, sheetIndex, sheetIndex);
    }

    private StubRefEval(
        int rowIndex,
        int columnIndex,
        ValueEval valueEval,
        int firstSheetIndex,
        int lastSheetIndex) {
      super(firstSheetIndex, lastSheetIndex, rowIndex, columnIndex);
      this.valueEval = valueEval;
    }

    @Override
    public ValueEval getInnerValueEval(int sheetIndex) {
      return valueEval;
    }

    @Override
    public AreaEval offset(
        int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx) {
      int firstRow = getRow() + relFirstRowIx;
      int lastRow = getRow() + relLastRowIx;
      int firstColumn = getColumn() + relFirstColIx;
      int lastColumn = getColumn() + relLastColIx;
      int height = lastRow - firstRow + 1;
      int width = lastColumn - firstColumn + 1;
      ValueEval[][] values = new ValueEval[height][width];
      for (int rowOffset = 0; rowOffset < height; rowOffset++) {
        for (int columnOffset = 0; columnOffset < width; columnOffset++) {
          values[rowOffset][columnOffset] = valueEval;
        }
      }
      return new StubAreaEval(firstRow, lastRow, firstColumn, lastColumn, values);
    }
  }

  /** Test-only `AreaEval` stub that returns pre-seeded values without POI workbook state. */
  private static final class StubAreaEval extends org.apache.poi.ss.formula.eval.AreaEvalBase {
    private final ValueEval[][] values;

    private StubAreaEval(
        int firstRow, int lastRow, int firstColumn, int lastColumn, ValueEval[][] values) {
      super(firstRow, firstColumn, lastRow, lastColumn);
      this.values = values;
    }

    @Override
    public ValueEval getRelativeValue(int relativeRowIndex, int relativeColumnIndex) {
      return values[relativeRowIndex][relativeColumnIndex];
    }

    @Override
    public ValueEval getRelativeValue(
        int sheetIndex, int relativeRowIndex, int relativeColumnIndex) {
      return getRelativeValue(relativeRowIndex, relativeColumnIndex);
    }

    @Override
    public AreaEval offset(
        int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx) {
      return this;
    }

    @Override
    public org.apache.poi.ss.formula.TwoDEval getRow(int rowIndex) {
      return this;
    }

    @Override
    public org.apache.poi.ss.formula.TwoDEval getColumn(int columnIndex) {
      return this;
    }
  }

  /** Test-only workbook that can simulate close failures without touching disk. */
  private static final class ThrowingWorkbook extends XSSFWorkbook {
    private final String message;
    private boolean failOnClose = true;

    private ThrowingWorkbook(String message) {
      this.message = message;
    }

    private void disableCloseFailure() {
      failOnClose = false;
    }

    @Override
    public void close() throws IOException {
      if (failOnClose) {
        throw new IOException(message);
      }
      super.close();
    }
  }
}
