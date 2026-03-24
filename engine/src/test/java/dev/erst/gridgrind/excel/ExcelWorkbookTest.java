package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.UUID;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Integration tests for ExcelWorkbook creation, loading, and sheet access. */
class ExcelWorkbookTest {
  @Test
  void snapshotsAndPreviewExposeFormulaResults() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-engine-", ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor commandExecutor = new WorkbookCommandExecutor();
      commandExecutor.apply(
          workbook,
          List.of(
              new WorkbookCommand.CreateSheet("Budget"),
              new WorkbookCommand.AppendRow(
                  "Budget", List.of(ExcelCellValue.text("Item"), ExcelCellValue.text("Amount"))),
              new WorkbookCommand.AppendRow(
                  "Budget", List.of(ExcelCellValue.text("Hosting"), ExcelCellValue.number(49.0))),
              new WorkbookCommand.AppendRow(
                  "Budget", List.of(ExcelCellValue.text("Domain"), ExcelCellValue.number(12.0))),
              new WorkbookCommand.SetCell("Budget", "A4", ExcelCellValue.text("Total")),
              new WorkbookCommand.SetCell("Budget", "B4", ExcelCellValue.formula("SUM(B2:B3)")),
              new WorkbookCommand.EvaluateAllFormulas()));
      workbook.save(workbookPath);
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = workbook.sheet("Budget");

      ExcelCellSnapshot.FormulaSnapshot totalSnapshot = (ExcelCellSnapshot.FormulaSnapshot) sheet.snapshotCell("B4");
      assertEquals("FORMULA", totalSnapshot.declaredType());
      assertEquals("FORMULA", totalSnapshot.effectiveType());
      assertEquals("SUM(B2:B3)", totalSnapshot.formula());
      assertEquals(61.0, ((ExcelCellSnapshot.NumberSnapshot)totalSnapshot.evaluation()).numberValue());

      List<ExcelPreviewRow> preview = sheet.preview(4, 2);
      assertEquals(4, preview.size());
      assertEquals("A1", preview.get(0).cells().get(0).address());
      assertEquals("Hosting", ((ExcelCellSnapshot.TextSnapshot) preview.get(1).cells().get(0)).stringValue());
      assertEquals("61", preview.get(3).cells().get(1).displayValue());
    }
  }

  @Test
  void managesWorkbookLifecycleAndValidation() throws IOException {
    Path workbookPath = Files.createTempDirectory("gridgrind-workbook-").resolve("nested/book.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Hello"));
      workbook.getOrCreateSheet("Budget").setCell("B1", ExcelCellValue.number(12.0));
      workbook.forceFormulaRecalculationOnOpen();
      workbook.save(workbookPath);

      assertEquals(1, workbook.sheetCount());
      assertEquals(List.of("Budget"), workbook.sheetNames());
      assertTrue(workbook.forceFormulaRecalculationOnOpenEnabled());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertEquals("Hello", workbook.sheet("Budget").text("A1"));
      assertEquals(12.0, workbook.sheet("Budget").number("B1"));
    }
  }

  @Test
  void wrapsWorkbookWideFormulaEvaluationFailures() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook,
                failingEvaluator(poiWorkbook.getCreationHelper().createFormulaEvaluator()))) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.formula("1+1"));

      InvalidFormulaException exception =
          assertThrows(InvalidFormulaException.class, workbook::evaluateAllFormulas);
      assertEquals("Budget", exception.sheetName());
      assertEquals("A1", exception.address());
      assertEquals("1+1", exception.formula());
    }
  }

  @Test
  void validatesWorkbookInputsAndMissingResources() throws IOException {
    assertThrows(NullPointerException.class, () -> ExcelWorkbook.open(null));

    Path missingPath = Files.createTempDirectory("gridgrind-missing-").resolve("missing.xlsx");
    WorkbookNotFoundException missingWorkbook =
        assertThrows(WorkbookNotFoundException.class, () -> ExcelWorkbook.open(missingPath));
    assertTrue(missingWorkbook.getMessage().contains("Workbook does not exist"));
    assertEquals(missingPath.toAbsolutePath(), missingWorkbook.workbookPath());

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      assertThrows(NullPointerException.class, () -> workbook.getOrCreateSheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.getOrCreateSheet(" "));
      assertThrows(NullPointerException.class, () -> workbook.sheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.sheet(" "));
      SheetNotFoundException missingSheet =
          assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Missing"));
      assertEquals("Missing", missingSheet.sheetName());
      assertThrows(NullPointerException.class, () -> workbook.save(null));
    }
  }

  @Test
  void savesToPathsWithoutParentDirectories() throws IOException {
    Path relativePath = Path.of("gridgrind-relative-" + UUID.randomUUID() + ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Hello"));
      workbook.save(relativePath);
      assertTrue(Files.exists(relativePath));
    } finally {
      Files.deleteIfExists(relativePath);
    }
  }

  private FormulaEvaluator failingEvaluator(FormulaEvaluator delegate) {
    InvocationHandler handler =
        new InvocationHandler() {
          @Override
          public Object invoke(Object proxy, java.lang.reflect.Method method, Object[] args)
              throws Throwable {
            if (method.getDeclaringClass() == Object.class) {
              return switch (method.getName()) {
                case "toString" -> "failingEvaluator";
                case "hashCode" -> System.identityHashCode(proxy);
                case "equals" -> sameProxyHandler(args, this);
                default -> throw new UnsupportedOperationException(method.getName());
              };
            }
            if ("evaluate".equals(method.getName()) || "evaluateAll".equals(method.getName())) {
              throw new org.apache.poi.ss.formula.FakeFormulaFailure("bad formula");
            }
            return method.invoke(delegate, args);
          }
        };
    return (FormulaEvaluator)
        Proxy.newProxyInstance(
            formulaEvaluatorClassLoader(), new Class<?>[] {FormulaEvaluator.class}, handler);
  }

  private ClassLoader formulaEvaluatorClassLoader() {
    return Objects.requireNonNull(
        Thread.currentThread().getContextClassLoader(), "context class loader must not be null");
  }

  private boolean sameProxyHandler(Object[] args, InvocationHandler handler) {
    if (args == null
        || args.length != 1
        || args[0] == null
        || !Proxy.isProxyClass(args[0].getClass())) {
      return false;
    }
    return Objects.equals(Proxy.getInvocationHandler(args[0]), handler);
  }
}
