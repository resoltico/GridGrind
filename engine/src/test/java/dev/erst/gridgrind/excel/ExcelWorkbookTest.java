package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.UUID;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddress;
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

      ExcelCellSnapshot.FormulaSnapshot totalSnapshot =
          (ExcelCellSnapshot.FormulaSnapshot) sheet.snapshotCell("B4");
      assertEquals("FORMULA", totalSnapshot.declaredType());
      assertEquals("FORMULA", totalSnapshot.effectiveType());
      assertEquals("SUM(B2:B3)", totalSnapshot.formula());
      assertEquals(
          61.0, ((ExcelCellSnapshot.NumberSnapshot) totalSnapshot.evaluation()).numberValue());

      List<ExcelPreviewRow> preview = sheet.preview(4, 2);
      assertEquals(4, preview.size());
      assertEquals("A1", preview.get(0).cells().get(0).address());
      assertEquals(
          "Hosting",
          ((ExcelCellSnapshot.TextSnapshot) preview.get(1).cells().get(0)).stringValue());
      assertEquals("61", preview.get(3).cells().get(1).displayValue());
    }
  }

  @Test
  void managesWorkbookLifecycleAndValidation() throws IOException {
    Path workbookPath =
        Files.createTempDirectory("gridgrind-workbook-").resolve("nested/book.xlsx");

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
  void renamesDeletesAndMovesSheetsAcrossSaves() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-sheet-management-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Live"));
      workbook.getOrCreateSheet("Archive").setCell("A1", ExcelCellValue.text("Old"));
      workbook.getOrCreateSheet("Scratch");

      workbook.renameSheet("Archive", "History");
      workbook.moveSheet("History", 0);
      workbook.deleteSheet("Scratch");
      workbook.save(workbookPath);

      assertEquals(List.of("History", "Budget"), workbook.sheetNames());
      assertEquals("Old", workbook.sheet("History").text("A1"));
      assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Archive"));
    }

    assertEquals(List.of("History", "Budget"), XlsxRoundTrip.sheetOrder(workbookPath));
  }

  @Test
  void persistsStructuralLayoutOperationsAcrossSaves() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Quarterly"));
      budget.mergeCells("A1:B1");
      budget.setColumnWidth(0, 1, 16.0);
      budget.setRowHeight(0, 0, 28.5);
      budget.freezePanes(1, 1, 1, 1);
      workbook.save(workbookPath);
    }

    assertEquals(List.of("A1:B1"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(
        new XlsxRoundTrip.FreezePaneState.Frozen(1, 1, 1, 1),
        XlsxRoundTrip.freezePaneState(workbookPath, "Budget"));
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
      workbook.getOrCreateSheet("Budget");
      workbook.getOrCreateSheet("Archive");
      assertSame(workbook, workbook.renameSheet("Budget", "Budget"));
      assertThrows(NullPointerException.class, () -> workbook.renameSheet(null, "Summary"));
      assertThrows(IllegalArgumentException.class, () -> workbook.renameSheet(" ", "Summary"));
      assertThrows(NullPointerException.class, () -> workbook.renameSheet("Budget", null));
      assertThrows(IllegalArgumentException.class, () -> workbook.renameSheet("Budget", " "));
      assertThrows(SheetNotFoundException.class, () -> workbook.renameSheet("Missing", "Summary"));
      assertThrows(IllegalArgumentException.class, () -> workbook.renameSheet("Budget", "Archive"));
      assertThrows(
          IllegalArgumentException.class, () -> workbook.renameSheet("Budget", "Bad/Name"));
      assertThrows(NullPointerException.class, () -> workbook.deleteSheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.deleteSheet(" "));
      assertThrows(SheetNotFoundException.class, () -> workbook.deleteSheet("Missing"));
      assertThrows(NullPointerException.class, () -> workbook.moveSheet(null, 0));
      assertThrows(IllegalArgumentException.class, () -> workbook.moveSheet(" ", 0));
      assertThrows(SheetNotFoundException.class, () -> workbook.moveSheet("Missing", 0));
      assertThrows(IllegalArgumentException.class, () -> workbook.moveSheet("Budget", -1));
      assertThrows(IllegalArgumentException.class, () -> workbook.moveSheet("Budget", 2));
      assertThrows(NullPointerException.class, () -> workbook.save(null));
    }
  }

  @Test
  void rejectsNonXlsxWorkbookFiles() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-legacy-", ".xls");

    try (HSSFWorkbook workbook = new HSSFWorkbook();
        var outputStream = Files.newOutputStream(workbookPath)) {
      workbook.createSheet("Budget").createRow(0).createCell(0).setCellValue("Legacy");
      workbook.write(outputStream);
    }

    IllegalArgumentException exception =
        assertThrows(IllegalArgumentException.class, () -> ExcelWorkbook.open(workbookPath));
    assertEquals("Only .xlsx workbooks are supported", exception.getMessage());
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

  @Test
  void roundTripHelpersInspectSavedWorkbookStructure() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-structure-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var budget = workbook.createSheet("Budget");
      workbook.createSheet("Summary");

      budget.addMergedRegion(CellRangeAddress.valueOf("A1:B2"));
      budget.setColumnWidth(0, 4096);
      budget.createRow(0).setHeightInPoints(28.5f);
      budget.createFreezePane(1, 2, 3, 4);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    assertEquals(List.of("Budget", "Summary"), XlsxRoundTrip.sheetOrder(workbookPath));
    assertEquals(List.of("A1:B2"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(
        new XlsxRoundTrip.FreezePaneState.Frozen(1, 2, 3, 4),
        XlsxRoundTrip.freezePaneState(workbookPath, "Budget"));
  }

  @Test
  void savesAndReopensFormattingDepthStyles() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-style-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Item"));
      workbook
          .sheet("Budget")
          .applyStyle(
              "A1",
              new ExcelCellStyle(
                  null,
                  true,
                  false,
                  true,
                  ExcelHorizontalAlignment.CENTER,
                  ExcelVerticalAlignment.TOP,
                  "Aptos",
                  ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
                  "#1F4E78",
                  true,
                  true,
                  "#FFF2CC",
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN),
                      null,
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE),
                      null,
                      null)));
      workbook.save(workbookPath);
    }

    ExcelCellStyleSnapshot style = XlsxRoundTrip.cellStyle(workbookPath, "Budget", "A1");
    assertTrue(style.bold());
    assertFalse(style.italic());
    assertTrue(style.wrapText());
    assertEquals(ExcelHorizontalAlignment.CENTER, style.horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, style.verticalAlignment());
    assertEquals("Aptos", style.fontName());
    assertEquals(new BigDecimal("11.5"), style.fontHeight().points());
    assertEquals("#1F4E78", style.fontColor());
    assertTrue(style.underline());
    assertTrue(style.strikeout());
    assertEquals("#FFF2CC", style.fillColor());
    assertEquals(ExcelBorderStyle.THIN, style.topBorderStyle());
    assertEquals(ExcelBorderStyle.DOUBLE, style.rightBorderStyle());
    assertEquals(ExcelBorderStyle.THIN, style.bottomBorderStyle());
    assertEquals(ExcelBorderStyle.THIN, style.leftBorderStyle());
  }

  @Test
  void persistsHyperlinksCommentsAndNamedRangesAcrossSaves() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-authoring-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Report"));
      budget.setCell("B4", ExcelCellValue.number(61.0));
      budget.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      budget.setComment("A1", new ExcelComment("Review", "GridGrind", true));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B4")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "LocalItem",
              new ExcelNamedRangeScope.SheetScope("Budget"),
              new ExcelNamedRangeTarget("Budget", "A1:B2")));
      workbook.save(workbookPath);
    }

    ExcelCellMetadataSnapshot metadata = XlsxRoundTrip.cellMetadata(workbookPath, "Budget", "A1");
    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"), metadata.hyperlink().orElseThrow());
    assertEquals(new ExcelComment("Review", "GridGrind", true), metadata.comment().orElseThrow());
    List<ExcelNamedRangeSnapshot> namedRanges = XlsxRoundTrip.namedRanges(workbookPath);
    assertEquals(2, namedRanges.size());
    assertTrue(
        namedRanges.contains(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                namedRanges.stream()
                    .filter(namedRange -> "BudgetTotal".equals(namedRange.name()))
                    .findFirst()
                    .orElseThrow()
                    .refersToFormula(),
                new ExcelNamedRangeTarget("Budget", "B4"))));
    assertTrue(
        namedRanges.contains(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "LocalItem",
                new ExcelNamedRangeScope.SheetScope("Budget"),
                namedRanges.stream()
                    .filter(namedRange -> "LocalItem".equals(namedRange.name()))
                    .findFirst()
                    .orElseThrow()
                    .refersToFormula(),
                new ExcelNamedRangeTarget("Budget", "A1:B2"))));

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertEquals(2, workbook.namedRangeCount());
      assertEquals(2, workbook.namedRanges().size());
      workbook.deleteNamedRange("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
      assertEquals(1, workbook.namedRangeCount());
    }
  }

  @Test
  void persistsLatestHyperlinkTargetAfterRepeatedWrites() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-hyperlink-replace-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("C");
      sheet.setHyperlink("F18", new ExcelHyperlink.Email("Report_Value@example.com"));
      sheet.setHyperlink("F18", new ExcelHyperlink.Email("Summary.Total@example.com"));
      workbook.save(workbookPath);
    }

    ExcelCellMetadataSnapshot metadata = XlsxRoundTrip.cellMetadata(workbookPath, "C", "F18");
    assertEquals(
        new ExcelHyperlink.Email("Summary.Total@example.com"), metadata.hyperlink().orElseThrow());
  }

  @Test
  void replacesNamedRangesAndExposesOnlySupportedNamedRangeSnapshots() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook, poiWorkbook.getCreationHelper().createFormulaEvaluator())) {
      poiWorkbook.createSheet("Budget");

      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B4")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "C1")));

      Name formulaName = poiWorkbook.createName();
      formulaName.setNameName("BudgetRollup");
      formulaName.setRefersToFormula("SUM(Budget!$B$2:$B$3)");
      Name internalLowerDefinedName = poiWorkbook.createName();
      internalLowerDefinedName.setNameName("_xlnm.Print_Area");
      internalLowerDefinedName.setRefersToFormula("Budget!$A$1");
      Name internalUpperDefinedName = poiWorkbook.createName();
      internalUpperDefinedName.setNameName("_XLNM.PRINT_TITLES");
      internalUpperDefinedName.setRefersToFormula("Budget!$A$1");

      Name workbookScoped = poiWorkbook.getName("BudgetTotal");
      Name hidden = syntheticName("HiddenBudgetTotal", false, true);
      Name internalLower = syntheticName("_xlnm.Print_Area", false, false);
      Name internalUpper = syntheticName("_XLNM.PRINT_TITLES", false, false);
      Name function = syntheticName("BudgetFn", true, false);
      Name sheetScoped = poiWorkbook.createName();
      sheetScoped.setNameName("LocalItem");
      sheetScoped.setSheetIndex(poiWorkbook.getSheetIndex("Budget"));
      sheetScoped.setRefersToFormula("Budget!$A$1");

      assertTrue(ExcelWorkbook.shouldExpose(workbookScoped));
      assertFalse(ExcelWorkbook.shouldExpose(syntheticName(null, false, false)));
      assertFalse(ExcelWorkbook.shouldExpose(hidden));
      assertFalse(ExcelWorkbook.shouldExpose(internalLower));
      assertFalse(ExcelWorkbook.shouldExpose(internalUpper));
      assertFalse(ExcelWorkbook.shouldExpose(function));

      assertTrue(workbook.scopeMatches(workbookScoped, new ExcelNamedRangeScope.WorkbookScope()));
      assertFalse(
          workbook.scopeMatches(workbookScoped, new ExcelNamedRangeScope.SheetScope("Budget")));
      assertTrue(workbook.scopeMatches(sheetScoped, new ExcelNamedRangeScope.SheetScope("Budget")));
      assertFalse(workbook.scopeMatches(sheetScoped, new ExcelNamedRangeScope.WorkbookScope()));
      assertThrows(
          SheetNotFoundException.class,
          () -> workbook.scopeMatches(sheetScoped, new ExcelNamedRangeScope.SheetScope("Missing")));

      assertEquals(
          List.of(
              new ExcelNamedRangeSnapshot.RangeSnapshot(
                  "BudgetTotal",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  "Budget!$C$1",
                  new ExcelNamedRangeTarget("Budget", "C1")),
              new ExcelNamedRangeSnapshot.FormulaSnapshot(
                  "BudgetRollup",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  "SUM(Budget!$B$2:$B$3)"),
              new ExcelNamedRangeSnapshot.RangeSnapshot(
                  "LocalItem",
                  new ExcelNamedRangeScope.SheetScope("Budget"),
                  "Budget!$A$1",
                  new ExcelNamedRangeTarget("Budget", "A1"))),
          workbook.namedRanges());
    }
  }

  @Test
  void validatesNamedRangeOperations() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");

      assertThrows(NullPointerException.class, () -> workbook.setNamedRange(null));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              workbook.setNamedRange(
                  new ExcelNamedRangeDefinition(
                      "A1",
                      new ExcelNamedRangeScope.WorkbookScope(),
                      new ExcelNamedRangeTarget("Budget", "B4"))));
      assertThrows(
          NullPointerException.class,
          () -> workbook.deleteNamedRange(null, new ExcelNamedRangeScope.WorkbookScope()));
      assertThrows(
          NullPointerException.class, () -> workbook.deleteNamedRange("BudgetTotal", null));
      assertThrows(
          NamedRangeNotFoundException.class,
          () -> workbook.deleteNamedRange("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              workbook.setNamedRange(
                  new ExcelNamedRangeDefinition(
                      "BudgetTotal",
                      new ExcelNamedRangeScope.SheetScope("Missing"),
                      new ExcelNamedRangeTarget("Missing", "A1"))));
    }
  }

  @Test
  void rejectsSheetNamesExceeding31Characters() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      String exactly31 = "A".repeat(31);
      String tooLong = "A".repeat(32);

      assertDoesNotThrow(() -> workbook.getOrCreateSheet(exactly31));
      assertThrows(IllegalArgumentException.class, () -> workbook.getOrCreateSheet(tooLong));
      assertThrows(IllegalArgumentException.class, () -> workbook.sheet(tooLong));
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

  private Name syntheticName(String nameName, boolean functionName, boolean hidden) {
    InvocationHandler handler =
        new InvocationHandler() {
          @Override
          public Object invoke(Object proxy, java.lang.reflect.Method method, Object[] args) {
            return switch (method.getName()) {
              case "getNameName" -> nameName;
              case "isFunctionName" -> functionName;
              case "isHidden" -> hidden;
              case "toString" -> "syntheticName[" + nameName + "]";
              case "hashCode" -> Objects.hash(nameName, functionName, hidden);
              case "equals" -> false;
              default -> throw new UnsupportedOperationException(method.getName());
            };
          }
        };
    return (Name)
        Proxy.newProxyInstance(formulaEvaluatorClassLoader(), new Class<?>[] {Name.class}, handler);
  }
}
