package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.formulaCell;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.textCell;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import org.junit.jupiter.api.Test;

/** Direct coverage for semantic selector resolution beyond end-to-end executor workflows. */
class SemanticSelectorResolverTest {
  private final SemanticSelectorResolver resolver =
      new SemanticSelectorResolver(new dev.erst.gridgrind.excel.WorkbookReadExecutor());

  @Test
  void resolvesSemanticExactCellTargetsAcrossSupportedKeyKinds() throws IOException {
    try (ExcelWorkbook workbook = workbookWithSemanticTables()) {
      assertResolvedAddress(textAmountCellTarget(), workbook, "Texts", "B2");
      assertResolvedAddress(textAmountCellTargetOnSheet(), workbook, "Texts", "B2");
      assertResolvedAddress(blankAmountCellTarget(), workbook, "Blanks", "B2");
      assertResolvedAddress(numberAmountCellTarget(), workbook, "Numbers", "B2");
      assertResolvedAddress(booleanAmountCellTarget(), workbook, "Flags", "B2");
      assertResolvedAddress(formulaAmountCellTarget(), workbook, "Formulas", "B2");
      assertResolvedAddress(wideAmountCellTarget(), workbook, "Wide", "AB2");

      CellSelector.ByAddress indexedRowTarget =
          assertInstanceOf(
              CellSelector.ByAddress.class,
              resolver.resolveMutationTarget(
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByIndex(new TableSelector.ByName("TextTable"), 0),
                      "Amount"),
                  new MutationAction.SetCell(textCell("Updated"))));
      assertEquals("Texts", indexedRowTarget.sheetName());
      assertEquals("B2", indexedRowTarget.address());
    }
  }

  @Test
  void returnsDirectTargetsWhenResolutionIsNotRequired() throws IOException {
    try (ExcelWorkbook workbook = workbookWithSemanticTables()) {
      CellSelector.ByAddress directCell = new CellSelector.ByAddress("Budget", "A1");
      assertSame(
          directCell,
          resolver.resolveMutationTarget(
              workbook, directCell, new MutationAction.SetCell(textCell("Owner"))));

      TableCellSelector.ByColumnName tableCell = textAmountCellTarget();
      assertSame(
          tableCell,
          resolver.resolveMutationTarget(workbook, tableCell, new MutationAction.ClearRange()));
      assertSame(
          tableCell,
          resolver.resolveAssertionTarget(
              workbook,
              tableCell,
              new Assertion.Not(new Assertion.CellValue(new ExpectedCellValue.Text("Updated")))));

      SemanticSelectorResolver.ResolvedInspectionTarget inspectionTarget =
          resolver.resolveInspectionTarget(
              "inspect-table-summary",
              workbook,
              tableCell,
              new InspectionQuery.GetWorkbookSummary());
      assertFalse(inspectionTarget.isShortCircuit());
      assertSame(tableCell, inspectionTarget.selector());
    }
  }

  @Test
  void shortCircuitsZeroMatchExactCellInspectionsAcrossSupportedQueries() throws IOException {
    try (ExcelWorkbook workbook = workbookWithSemanticTables()) {
      TableCellSelector.ByColumnName missingTarget =
          new TableCellSelector.ByColumnName(
              new TableRowSelector.ByKeyCell(
                  new TableSelector.ByName("TextTable"), "Item", textCell("Missing")),
              "Amount");

      SemanticSelectorResolver.ResolvedInspectionTarget cellsResult =
          resolver.resolveInspectionTarget(
              "inspect-cells", workbook, missingTarget, new InspectionQuery.GetCells());
      assertTrue(cellsResult.isShortCircuit());
      assertEquals(
          "Texts",
          assertInstanceOf(InspectionResult.CellsResult.class, cellsResult.shortCircuitResult())
              .sheetName());

      SemanticSelectorResolver.ResolvedInspectionTarget hyperlinksResult =
          resolver.resolveInspectionTarget(
              "inspect-links", workbook, missingTarget, new InspectionQuery.GetHyperlinks());
      assertTrue(hyperlinksResult.isShortCircuit());
      assertEquals(
          "Texts",
          assertInstanceOf(
                  InspectionResult.HyperlinksResult.class, hyperlinksResult.shortCircuitResult())
              .sheetName());

      SemanticSelectorResolver.ResolvedInspectionTarget commentsResult =
          resolver.resolveInspectionTarget(
              "inspect-comments", workbook, missingTarget, new InspectionQuery.GetComments());
      assertTrue(commentsResult.isShortCircuit());
      assertEquals(
          "Texts",
          assertInstanceOf(
                  InspectionResult.CommentsResult.class, commentsResult.shortCircuitResult())
              .sheetName());
    }
  }

  @Test
  void keyedResolutionTreatsNonMatchingSupportedScalarKeysAsZeroMatch() throws IOException {
    try (ExcelWorkbook workbook = workbookWithSemanticTables()) {
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-text-type-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("NumberTable"), "Item", textCell("Hosting")),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-blank-type-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("TextTable"), "Item", new CellInput.Blank()),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-number-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("NumberTable"),
                          "Item",
                          new CellInput.Numeric(99.0d)),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-boolean-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("BooleanTable"),
                          "Item",
                          new CellInput.BooleanValue(false)),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-formula-type-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("TextTable"), "Item", formulaCell("2+2")),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-formula-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("FormulaTable"), "Item", formulaCell("1+1")),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-number-type-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("TextTable"),
                          "Item",
                          new CellInput.Numeric(42.0d)),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  "inspect-boolean-type-miss",
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName("TextTable"),
                          "Item",
                          new CellInput.BooleanValue(true)),
                      "Amount"),
                  new InspectionQuery.GetCells())
              .isShortCircuit());
    }
  }

  @Test
  void rejectsInvalidSemanticResolutionCasesInsteadOfGuessing() throws IOException {
    try (ExcelWorkbook workbook = workbookWithSemanticTables()) {
      IllegalArgumentException missingRow =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveMutationTarget(
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByName("TextTable"), "Item", textCell("Missing")),
                          "Amount"),
                      new MutationAction.SetCell(textCell("Updated"))));
      assertTrue(missingRow.getMessage().contains("matched no rows"));

      IllegalArgumentException outOfRange =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveAssertionTarget(
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByIndex(new TableSelector.ByName("TextTable"), 4),
                          "Amount"),
                      new Assertion.DisplayValue("100")));
      assertTrue(outOfRange.getMessage().contains("outside the data-row bounds"));

      IllegalArgumentException missingTable =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveInspectionTarget(
                      "inspect-missing-table",
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByName("MissingTable"),
                              "Item",
                              textCell("Hosting")),
                          "Amount"),
                      new InspectionQuery.GetCells()));
      assertTrue(missingTable.getMessage().contains("table not found"));

      IllegalArgumentException missingTableOnExpectedSheet =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveInspectionTarget(
                      "inspect-missing-table-on-sheet",
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByNameOnSheet("TextTable", "Budget"),
                              "Item",
                              textCell("Hosting")),
                          "Amount"),
                      new InspectionQuery.GetCells()));
      assertTrue(
          missingTableOnExpectedSheet.getMessage().contains("table not found on expected sheet"));

      IllegalArgumentException missingNamedTableOnExpectedSheet =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveInspectionTarget(
                      "inspect-missing-named-table-on-sheet",
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByNameOnSheet("MissingTable", "Budget"),
                              "Item",
                              textCell("Hosting")),
                          "Amount"),
                      new InspectionQuery.GetCells()));
      assertTrue(
          missingNamedTableOnExpectedSheet
              .getMessage()
              .contains("table not found on expected sheet: MissingTable@Budget"));

      IllegalArgumentException missingColumn =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveAssertionTarget(
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByName("TextTable"),
                              "MissingColumn",
                              textCell("Hosting")),
                          "Amount"),
                      new Assertion.DisplayValue("100")));
      assertTrue(missingColumn.getMessage().contains("does not contain column MissingColumn"));

      IllegalStateException unresolvedTextSource =
          assertThrows(
              IllegalStateException.class,
              () ->
                  resolver.resolveMutationTarget(
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByName("TextTable"),
                              "Item",
                              new CellInput.Text(TextSourceInput.utf8File("item.txt"))),
                          "Amount"),
                      new MutationAction.SetCell(textCell("Updated"))));
      assertTrue(unresolvedTextSource.getMessage().contains("must be resolved to INLINE text"));
    }
  }

  @Test
  void resolvedInspectionTargetRejectsNullSelectorAndNullShortCircuit() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new SemanticSelectorResolver.ResolvedInspectionTarget(null, null));
    assertTrue(
        failure
            .getMessage()
            .contains(
                "resolved inspection target requires either a selector or a short-circuit result"));
  }

  private void assertResolvedAddress(
      TableCellSelector.ByColumnName target,
      ExcelWorkbook workbook,
      String expectedSheet,
      String expectedAddress) {
    CellSelector.ByAddress mutationTarget =
        assertInstanceOf(
            CellSelector.ByAddress.class,
            resolver.resolveMutationTarget(
                workbook, target, new MutationAction.SetCell(textCell("Updated"))));
    assertEquals(expectedSheet, mutationTarget.sheetName());
    assertEquals(expectedAddress, mutationTarget.address());

    CellSelector.ByAddress assertionTarget =
        assertInstanceOf(
            CellSelector.ByAddress.class,
            resolver.resolveAssertionTarget(workbook, target, new Assertion.DisplayValue("100")));
    assertEquals(expectedSheet, assertionTarget.sheetName());
    assertEquals(expectedAddress, assertionTarget.address());

    GridGrindResponse.CellStyleReport expectedStyle =
        InspectionResultCellReportSupport.toCellReport(
                workbook.sheet(expectedSheet).snapshotCell(expectedAddress))
            .style();
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(
                    workbook,
                    target,
                    new MutationAction.SetHyperlink(
                        new HyperlinkTarget.Url("https://example.com"))))
            .address());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(
                    workbook, target, new MutationAction.ClearHyperlink()))
            .address());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(
                    workbook,
                    target,
                    new MutationAction.SetComment(
                        new CommentInput(
                            TextSourceInput.inline("Note"),
                            "Ada",
                            false,
                            java.util.Optional.empty(),
                            java.util.Optional.empty()))))
            .address());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(workbook, target, new MutationAction.ClearComment()))
            .address());

    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveAssertionTarget(
                    workbook,
                    target,
                    new Assertion.CellValue(new ExpectedCellValue.NumericValue(100.0d))))
            .address());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveAssertionTarget(workbook, target, new Assertion.FormulaText("2+2")))
            .address());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveAssertionTarget(
                    workbook, target, new Assertion.CellStyle(expectedStyle)))
            .address());

    SemanticSelectorResolver.ResolvedInspectionTarget inspectionTarget =
        resolver.resolveInspectionTarget(
            "inspect-cell", workbook, target, new InspectionQuery.GetCells());
    assertFalse(inspectionTarget.isShortCircuit());
    CellSelector.ByAddress inspectionAddress =
        assertInstanceOf(CellSelector.ByAddress.class, inspectionTarget.selector());
    assertEquals(expectedSheet, inspectionAddress.sheetName());
    assertEquals(expectedAddress, inspectionAddress.address());

    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver
                    .resolveInspectionTarget(
                        "inspect-links", workbook, target, new InspectionQuery.GetHyperlinks())
                    .selector())
            .address());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver
                    .resolveInspectionTarget(
                        "inspect-comments", workbook, target, new InspectionQuery.GetComments())
                    .selector())
            .address());
  }

  private static TableCellSelector.ByColumnName textAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("TextTable"), "Item", textCell("Hosting")),
        "Amount");
  }

  private static TableCellSelector.ByColumnName textAmountCellTargetOnSheet() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByNameOnSheet("TextTable", "Texts"), "Item", textCell("Hosting")),
        "Amount");
  }

  private static TableCellSelector.ByColumnName blankAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("BlankTable"), "Item", new CellInput.Blank()),
        "Amount");
  }

  private static TableCellSelector.ByColumnName numberAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("NumberTable"), "Item", new CellInput.Numeric(42.0d)),
        "Amount");
  }

  private static TableCellSelector.ByColumnName booleanAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("BooleanTable"), "Item", new CellInput.BooleanValue(true)),
        "Amount");
  }

  private static TableCellSelector.ByColumnName formulaAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("FormulaTable"), "Item", formulaCell("2+2")),
        "Amount");
  }

  private static TableCellSelector.ByColumnName wideAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("WideTable"), "Item", textCell("Hosting")),
        "Amount");
  }

  private static ExcelWorkbook workbookWithSemanticTables() throws IOException {
    ExcelWorkbook workbook = ExcelWorkbook.create();
    addTable(
        workbook,
        "Texts",
        "TextTable",
        ExcelCellValue.text("Hosting"),
        ExcelCellValue.number(100.0d));
    addTable(
        workbook, "Blanks", "BlankTable", ExcelCellValue.blank(), ExcelCellValue.number(10.0d));
    addTable(
        workbook,
        "Numbers",
        "NumberTable",
        ExcelCellValue.number(42.0d),
        ExcelCellValue.number(200.0d));
    addTable(
        workbook,
        "Flags",
        "BooleanTable",
        ExcelCellValue.bool(true),
        ExcelCellValue.number(300.0d));
    addTable(
        workbook,
        "Formulas",
        "FormulaTable",
        ExcelCellValue.formula("2+2"),
        ExcelCellValue.number(400.0d));
    addWideTable(workbook);
    workbook.getOrCreateSheet("Budget");
    workbook.sheet("Budget").setCell("A1", ExcelCellValue.text("Owner"));
    return workbook;
  }

  private static void addTable(
      ExcelWorkbook workbook,
      String sheetName,
      String tableName,
      ExcelCellValue keyValue,
      ExcelCellValue amountValue) {
    workbook.getOrCreateSheet(sheetName);
    workbook.sheet(sheetName).setCell("A1", ExcelCellValue.text("Item"));
    workbook.sheet(sheetName).setCell("B1", ExcelCellValue.text("Amount"));
    workbook.sheet(sheetName).setCell("A2", keyValue);
    workbook.sheet(sheetName).setCell("B2", amountValue);
    workbook.setTable(
        new ExcelTableDefinition(tableName, sheetName, "A1:B2", false, new ExcelTableStyle.None()));
  }

  private static void addWideTable(ExcelWorkbook workbook) {
    workbook.getOrCreateSheet("Wide");
    workbook.sheet("Wide").setCell("AA1", ExcelCellValue.text("Item"));
    workbook.sheet("Wide").setCell("AB1", ExcelCellValue.text("Amount"));
    workbook.sheet("Wide").setCell("AA2", ExcelCellValue.text("Hosting"));
    workbook.sheet("Wide").setCell("AB2", ExcelCellValue.number(500.0d));
    workbook.setTable(
        new ExcelTableDefinition(
            "WideTable", "Wide", "AA1:AB2", false, new ExcelTableStyle.None()));
  }
}
