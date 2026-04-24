package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.XlsxRoundTrip;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Test;

/** Style, formula, and persistence-path tests for DefaultGridGrindRequestExecutor. */
class DefaultGridGrindRequestExecutorStyleAndFormulaTest
    extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void executesRangeAndStyleWorkflowAndSurfacesStyledCells() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-range-style-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:B2"),
                                new MutationAction.SetRange(
                                    List.of(
                                        List.of(textCell("Item"), textCell("Amount")),
                                        List.of(
                                            textCell("Hosting"), new CellInput.Numeric(49.0))))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:B1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        "#,##0.00",
                                        new CellAlignmentInput(
                                            true,
                                            ExcelHorizontalAlignment.CENTER,
                                            ExcelVerticalAlignment.CENTER,
                                            null,
                                            null),
                                        new CellFontInput(true, null, null, null, null, null, null),
                                        null,
                                        null,
                                        null))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "C1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        new CellAlignmentInput(
                                            null,
                                            ExcelHorizontalAlignment.RIGHT,
                                            ExcelVerticalAlignment.BOTTOM,
                                            null,
                                            null),
                                        new CellFontInput(null, true, null, null, null, null, null),
                                        null,
                                        null,
                                        null))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "B3"),
                                new MutationAction.SetCell(formulaCell("SUM(B2:B2)"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A2"),
                                new MutationAction.ClearRange())),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A1", "A2", "B3", "C1")),
                            new InspectionQuery.GetCells()),
                        inspect(
                            "window",
                            new RangeSelector.RectangularWindow("Budget", "A1", 3, 3),
                            new InspectionQuery.GetWindow()))));

    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.WindowReport window =
        read(success, "window", InspectionResult.WindowResult.class).window();

    assertTrue(Files.exists(workbookPath));
    assertEquals(
        ExcelHorizontalAlignment.CENTER,
        cells.cells().getFirst().style().alignment().horizontalAlignment());
    assertTrue(cells.cells().getFirst().style().font().bold());
    assertTrue(cells.cells().getFirst().style().alignment().wrapText());
    assertEquals("BLANK", cells.cells().get(1).effectiveType());
    assertEquals("49", cells.cells().get(2).displayValue());
    assertTrue(
        window.rows().getFirst().cells().stream().anyMatch(cell -> "C1".equals(cell.address())));
    assertTrue(cells.cells().get(3).style().font().italic());
  }

  @Test
  void executesFormattingDepthWorkflowAndPersistsReportedStyleState() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-format-depth-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(textCell("Item"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        new CellAlignmentInput(
                                            true,
                                            ExcelHorizontalAlignment.CENTER,
                                            ExcelVerticalAlignment.TOP,
                                            45,
                                            3),
                                        new CellFontInput(
                                            true,
                                            false,
                                            "Aptos",
                                            new FontHeightInput.Points(new BigDecimal("11.5")),
                                            "#1F4E78",
                                            true,
                                            true),
                                        new CellFillInput(
                                            dev.erst.gridgrind.excel.foundation.ExcelFillPattern
                                                .THIN_HORIZONTAL_BANDS,
                                            "#FFF2CC",
                                            "#DDEBF7"),
                                        new CellBorderInput(
                                            new CellBorderSideInput(
                                                ExcelBorderStyle.THIN, "#102030"),
                                            null,
                                            new CellBorderSideInput(
                                                ExcelBorderStyle.DOUBLE, "#203040"),
                                            null,
                                            null),
                                        new CellProtectionInput(false, true))))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A1")),
                            new InspectionQuery.GetCells()))));

    GridGrindResponse.CellStyleReport style =
        read(success, "cells", InspectionResult.CellsResult.class).cells().getFirst().style();

    assertTrue(Files.exists(workbookPath));
    assertTrue(style.font().bold());
    assertFalse(style.font().italic());
    assertTrue(style.alignment().wrapText());
    assertEquals(ExcelHorizontalAlignment.CENTER, style.alignment().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, style.alignment().verticalAlignment());
    assertEquals(45, style.alignment().textRotation());
    assertEquals(3, style.alignment().indentation());
    assertEquals("Aptos", style.font().fontName());
    assertEquals(230, style.font().fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), style.font().fontHeight().points());
    assertEquals(rgb("#1F4E78"), style.font().fontColor());
    assertTrue(style.font().underline());
    assertTrue(style.font().strikeout());
    assertEquals(
        dev.erst.gridgrind.excel.foundation.ExcelFillPattern.THIN_HORIZONTAL_BANDS,
        style.fill().pattern());
    assertEquals(rgb("#FFF2CC"), style.fill().foregroundColor());
    assertEquals(rgb("#DDEBF7"), style.fill().backgroundColor());
    assertEquals(ExcelBorderStyle.THIN, style.border().top().style());
    assertEquals(ExcelBorderStyle.DOUBLE, style.border().right().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().bottom().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().left().style());
    assertEquals(rgb("#102030"), style.border().top().color());
    assertEquals(rgb("#203040"), style.border().right().color());
    assertFalse(style.protection().locked());
    assertTrue(style.protection().hiddenFormula());
    assertEquals(
        style, toResponseStyleReport(XlsxRoundTrip.cellStyle(workbookPath, "Budget", "A1")));
  }

  @Test
  void executesAdvancedStyleWorkflowWithThemeIndexedAndGradientColors() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-advanced-style-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1:A2"),
                                new MutationAction.SetRange(
                                    List.of(
                                        List.of(textCell("ThemeTintStyle")),
                                        List.of(textCell("GradientFillStyle"))))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A1"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        new CellFontInput(
                                            null, true, null, null, null, 6, null, -0.35d, null,
                                            null),
                                        new CellFillInput(
                                            dev.erst.gridgrind.excel.foundation.ExcelFillPattern
                                                .SOLID,
                                            null,
                                            3,
                                            null,
                                            0.30d,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null),
                                        new CellBorderInput(
                                            null,
                                            null,
                                            null,
                                            new CellBorderSideInput(
                                                ExcelBorderStyle.THIN,
                                                null,
                                                null,
                                                Short.toUnsignedInt(
                                                    IndexedColors.DARK_RED.getIndex()),
                                                null),
                                            null),
                                        null))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A2"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        null,
                                        new CellFillInput(
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            new CellGradientFillInput(
                                                "LINEAR",
                                                45.0d,
                                                null,
                                                null,
                                                null,
                                                null,
                                                List.of(
                                                    new CellGradientStopInput(
                                                        0.0d, new ColorInput("#1F497D")),
                                                    new CellGradientStopInput(
                                                        1.0d,
                                                        new ColorInput(null, 4, null, 0.45d))))),
                                        null,
                                        new CellProtectionInput(true, true))))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A1", "A2")),
                            new InspectionQuery.GetCells()))));

    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.CellStyleReport themedStyle = cells.cells().get(0).style();
    GridGrindResponse.CellStyleReport gradientStyle = cells.cells().get(1).style();

    assertTrue(Files.exists(workbookPath));
    assertEquals(new CellColorReport(null, 6, null, -0.35d), themedStyle.font().fontColor());
    assertEquals(new CellColorReport(null, 3, null, 0.30d), themedStyle.fill().foregroundColor());
    assertEquals(
        new CellColorReport(
            null, null, Short.toUnsignedInt(IndexedColors.DARK_RED.getIndex()), null),
        themedStyle.border().bottom().color());
    assertNotNull(gradientStyle.fill().gradient());
    assertEquals("LINEAR", gradientStyle.fill().gradient().type());
    assertEquals(45.0d, gradientStyle.fill().gradient().degree());
    assertEquals(
        new CellColorReport("#1F497D", null, null, null),
        gradientStyle.fill().gradient().stops().get(0).color());
    assertEquals(
        new CellColorReport(null, 4, null, 0.45d),
        gradientStyle.fill().gradient().stops().get(1).color());
    assertTrue(gradientStyle.protection().locked());
    assertTrue(gradientStyle.protection().hiddenFormula());
  }

  @Test
  void preservesDistinctLinearAndPathGradientStylesInSameRequest() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-distinct-gradients-", ".xlsx");
    assertDoesNotThrow(() -> Files.deleteIfExists(workbookPath));
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A2"),
                                new MutationAction.SetCell(textCell("Linear gradient"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A2"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        null,
                                        new CellFillInput(
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            new CellGradientFillInput(
                                                "LINEAR",
                                                45.0d,
                                                null,
                                                null,
                                                null,
                                                null,
                                                List.of(
                                                    new CellGradientStopInput(
                                                        0.0d, new ColorInput("#1F497D")),
                                                    new CellGradientStopInput(
                                                        1.0d,
                                                        new ColorInput(null, 4, null, 0.45d))))),
                                        null,
                                        new CellProtectionInput(true, true)))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A3"),
                                new MutationAction.SetCell(textCell("Path gradient"))),
                            mutate(
                                new RangeSelector.ByRange("Budget", "A3"),
                                new MutationAction.ApplyStyle(
                                    new CellStyleInput(
                                        null,
                                        null,
                                        null,
                                        new CellFillInput(
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            new CellGradientFillInput(
                                                "PATH",
                                                null,
                                                0.1d,
                                                0.2d,
                                                0.3d,
                                                0.4d,
                                                List.of(
                                                    new CellGradientStopInput(
                                                        0.0d, new ColorInput("#112233")),
                                                    new CellGradientStopInput(
                                                        1.0d,
                                                        new ColorInput(
                                                            null,
                                                            null,
                                                            Short.toUnsignedInt(
                                                                IndexedColors.DARK_RED.getIndex()),
                                                            null))))),
                                        null,
                                        new CellProtectionInput(false, true))))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Budget", List.of("A2", "A3")),
                            new InspectionQuery.GetCells()))));

    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);
    GridGrindResponse.CellStyleReport linearGradientStyle = cells.cells().get(0).style();
    GridGrindResponse.CellStyleReport pathGradientStyle = cells.cells().get(1).style();

    assertEquals("LINEAR", linearGradientStyle.fill().gradient().type());
    assertEquals(45.0d, linearGradientStyle.fill().gradient().degree());
    assertTrue(linearGradientStyle.protection().locked());
    assertTrue(linearGradientStyle.protection().hiddenFormula());
    assertEquals("PATH", pathGradientStyle.fill().gradient().type());
    assertNull(pathGradientStyle.fill().gradient().degree());
    assertEquals(0.1d, pathGradientStyle.fill().gradient().left());
    assertEquals(0.2d, pathGradientStyle.fill().gradient().right());
    assertEquals(0.3d, pathGradientStyle.fill().gradient().top());
    assertEquals(0.4d, pathGradientStyle.fill().gradient().bottom());
    assertFalse(pathGradientStyle.protection().locked());
    assertTrue(pathGradientStyle.protection().hiddenFormula());
  }

  @Test
  void producesErrorReportForCellsWithErrorValues() {
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Data"), new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Data", "A1"),
                                new MutationAction.SetCell(formulaCell("1/0")))),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Data", List.of("A1")),
                            new InspectionQuery.GetCells()))));

    GridGrindResponse.CellReport cell =
        read(success, "cells", InspectionResult.CellsResult.class).cells().getFirst();
    assertInstanceOf(GridGrindResponse.CellReport.FormulaReport.class, cell);
    GridGrindResponse.CellReport evaluation =
        cast(GridGrindResponse.CellReport.FormulaReport.class, cell).evaluation();
    assertInstanceOf(GridGrindResponse.CellReport.ErrorReport.class, evaluation);
    assertEquals("ERROR", evaluation.effectiveType());
  }

  @Test
  void extractsFormulaFromSetCellOperationWhenExceptionCarriesNone() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation operation =
        mutate(
            new CellSelector.ByAddress("Data", "A1"),
            new MutationAction.SetCell(formulaCell("SUM(B1:B2)")));

    assertEquals("SUM(B1:B2)", formulaFor(operation, exception));
    assertEquals("Data", sheetNameFor(operation, exception));
    assertEquals("A1", addressFor(operation, exception));
    assertNull(rangeFor(operation, exception));
  }

  @Test
  void persistencePathResolvesCorrectlyForAllPersistenceAndSourceCombinations() {
    WorkbookPlan.WorkbookSource newSource = new WorkbookPlan.WorkbookSource.New();
    WorkbookPlan.WorkbookSource existingFile =
        new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/source.xlsx");
    WorkbookPlan.WorkbookPersistence none = new WorkbookPlan.WorkbookPersistence.None();
    WorkbookPlan.WorkbookPersistence overwrite =
        new WorkbookPlan.WorkbookPersistence.OverwriteSource();
    WorkbookPlan.WorkbookPersistence saveAs =
        new WorkbookPlan.WorkbookPersistence.SaveAs("/tmp/out.xlsx");

    assertEquals(
        Path.of("/tmp/out.xlsx").toAbsolutePath().normalize().toString(),
        ExecutionRequestPaths.persistencePath(newSource, saveAs));
    assertEquals(
        Path.of("/tmp/source.xlsx").toAbsolutePath().normalize().toString(),
        ExecutionRequestPaths.persistencePath(existingFile, overwrite));
    assertNull(ExecutionRequestPaths.persistencePath(newSource, overwrite));
    assertNull(ExecutionRequestPaths.persistencePath(newSource, none));
    assertNull(ExecutionRequestPaths.persistencePath(existingFile, none));
  }

  @Test
  void persistWorkbookRejectsOverwriteForNewSources() throws Exception {
    ExecutionWorkbookSupport workbookSupport = new ExecutionWorkbookSupport(Files::createTempFile);

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  workbookSupport.persistWorkbook(
                      workbook,
                      new WorkbookPlan.WorkbookSource.New(),
                      new WorkbookPlan.WorkbookPersistence.OverwriteSource()));

      assertEquals("OVERWRITE persistence requires an EXISTING source", exception.getMessage());
    }
  }

  @Test
  void persistencePathNormalizesDoubleDotSegments() {
    WorkbookPlan.WorkbookSource newSource = new WorkbookPlan.WorkbookSource.New();
    WorkbookPlan.WorkbookPersistence saveAs =
        new WorkbookPlan.WorkbookPersistence.SaveAs("/tmp/subdir/../out.xlsx");

    assertEquals("/tmp/out.xlsx", ExecutionRequestPaths.persistencePath(newSource, saveAs));
  }

  @Test
  void persistWorkbookSaveAsReportsNormalizedExecutionPath() throws Exception {
    ExecutionWorkbookSupport workbookSupport = new ExecutionWorkbookSupport(Files::createTempFile);
    Path tempDir = Files.createTempDirectory("gridgrind-normalize-test-");
    Path subDir = Files.createDirectory(tempDir.resolve("subdir"));
    String pathWithDotDot = subDir + "/../out.xlsx";

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      GridGrindResponse.PersistenceOutcome outcome =
          workbookSupport.persistWorkbook(
              workbook,
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.SaveAs(pathWithDotDot));

      GridGrindResponse.PersistenceOutcome.SavedAs savedAs =
          assertInstanceOf(GridGrindResponse.PersistenceOutcome.SavedAs.class, outcome);
      assertEquals(pathWithDotDot, savedAs.requestedPath());
      assertEquals(tempDir.resolve("out.xlsx").toString(), savedAs.executionPath());
    } finally {
      Files.deleteIfExists(tempDir.resolve("out.xlsx"));
      Files.deleteIfExists(subDir);
      Files.deleteIfExists(tempDir);
    }
  }

  @Test
  void guardsCatastrophicRuntimeExceptionsAndProducesExecuteRequestFailure() {
    int[] callCount = {0};
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new DefaultGridGrindRequestExecutorDependencies(
                new WorkbookCommandExecutor(),
                new WorkbookReadExecutor(),
                workbook -> {
                  int count = callCount[0];
                  callCount[0] = count + 1;
                  if (count == 0) {
                    throw new IllegalStateException("catastrophic close failure");
                  }
                },
                Files::createTempFile,
                dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close,
                dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter::markRecalculateOnOpen));

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void rejectsInvalidClearRangeSelectorsAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                mutate(
                    new RangeSelector.ByRange("Budget", "A1:"), new MutationAction.ClearRange()));

    assertEquals("range address must not be blank", failure.getMessage());
  }

  @Test
  void rejectsInvalidSetRangeSelectorsAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                mutate(
                    new RangeSelector.ByRange("Budget", "A1:"),
                    new MutationAction.SetRange(List.of(List.of(textCell("x"))))));

    assertEquals("range address must not be blank", failure.getMessage());
  }

  @Test
  void rejectsInvalidApplyStyleSelectorsAtContractConstructionTime() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                mutate(
                    new RangeSelector.ByRange("Budget", "A1:"),
                    new MutationAction.ApplyStyle(
                        new CellStyleInput(
                            null,
                            null,
                            new CellFontInput(true, null, null, null, null, null, null),
                            null,
                            null,
                            null))));

    assertEquals("range address must not be blank", failure.getMessage());
  }

  @Test
  void returnsStructuredFailureForAppendRowWithInvalidFormula() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.AppendRow(List.of(formulaCell("SUM("))))))));

    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("APPEND_ROW", failure.problem().context().stepType());
    assertEquals("Budget", failure.problem().context().sheetName());
  }

  @Test
  void extractsNullContextForOperationsWithNoSheetAddressRangeOrFormula() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation clearWorkbookProtection =
        mutate(new WorkbookSelector.Current(), new MutationAction.ClearWorkbookProtection());
    ExecutorTestPlanSupport.PendingMutation setWorkbookProtection =
        mutate(
            new WorkbookSelector.Current(),
            new MutationAction.SetWorkbookProtection(
                new WorkbookProtectionInput(true, false, false, null, null)));
    ExecutorTestPlanSupport.PendingMutation appendRow =
        mutate(
            new SheetSelector.ByName("Budget"),
            new MutationAction.AppendRow(List.of(textCell("x"))));
    ExecutorTestPlanSupport.PendingMutation ensureSheet =
        mutate(new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet());

    assertNull(formulaFor(clearWorkbookProtection, exception));
    assertNull(formulaFor(setWorkbookProtection, exception));
    assertNull(formulaFor(appendRow, exception));
    assertNull(formulaFor(ensureSheet, exception));
    assertNull(sheetNameFor(clearWorkbookProtection, exception));
    assertNull(sheetNameFor(setWorkbookProtection, exception));
    assertNull(addressFor(clearWorkbookProtection, exception));
    assertNull(addressFor(setWorkbookProtection, exception));
    assertNull(rangeFor(clearWorkbookProtection, exception));
    assertNull(rangeFor(setWorkbookProtection, exception));
  }

  @Test
  void extractsNullFormulaFromSetCellWithNonFormulaValueWhenExceptionCarriesNone() {
    RuntimeException exception = new RuntimeException("test");
    ExecutorTestPlanSupport.PendingMutation operation =
        mutate(
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetCell(textCell("hello")));

    assertNull(formulaFor(operation, exception));
    assertEquals("Budget", sheetNameFor(operation, exception));
    assertEquals("A1", addressFor(operation, exception));
    assertNull(rangeFor(operation, exception));
  }

  @Test
  void formulaForSetCellReturnsNullForAllNonFormulaValueTypes() {
    RuntimeException exception = new RuntimeException("test");

    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.Blank())),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(textCell("hello"))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(
                    new CellInput.RichText(
                        List.of(
                            richTextRun("Budget"),
                            new RichTextRunInput(
                                text(" FY26"),
                                new CellFontInput(
                                    true, false, null, null, "#AABBCC", false, false)))))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.Numeric(1.0))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.BooleanValue(true))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(new CellInput.Date(LocalDate.of(2024, 1, 1)))),
            exception));
    assertNull(
        formulaFor(
            mutate(
                new CellSelector.ByAddress("S", "A1"),
                new MutationAction.SetCell(
                    new CellInput.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0)))),
            exception));
  }
}
