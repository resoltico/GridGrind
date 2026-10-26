package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.formulaCell;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.textCell;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import java.io.IOException;
import org.junit.jupiter.api.Test;

/** Direct coverage for semantic selector resolution beyond end-to-end executor workflows. */
class SemanticSelectorResolverTest {
  private final SemanticSelectorResolver resolver =
      new SemanticSelectorResolver(new dev.erst.gridgrind.excel.WorkbookExecutionEngine());

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
                      new TableRowSelector.ByIndex(
                          new TableSelector.ByName(new TableName("TextTable")), 0),
                      "Amount"),
                  new CellMutationAction.SetCell(textCell("Updated"))));
      assertEquals("Texts", indexedRowTarget.sheetName().value());
      assertEquals("B2", indexedRowTarget.address().value());
    }
  }

  @Test
  void returnsDirectTargetsWhenResolutionIsNotRequired() throws IOException {
    try (ExcelWorkbook workbook = workbookWithSemanticTables()) {
      CellSelector.ByAddress directCell =
          new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"));
      assertSame(
          directCell,
          resolver.resolveMutationTarget(
              workbook, directCell, new CellMutationAction.SetCell(textCell("Owner"))));

      TableCellSelector.ByColumnName tableCell = textAmountCellTarget();
      assertSame(
          tableCell,
          resolver.resolveMutationTarget(workbook, tableCell, new CellMutationAction.ClearRange()));
      assertSame(
          tableCell,
          resolver.resolveAssertionTarget(
              workbook,
              tableCell,
              new CompositeAssertion.Not(
                  new CellAssertion.Value(
                      new dev.erst.gridgrind.core.model.value.CellValue.Text("Updated")))));

      SemanticSelectorResolver.ResolvedInspectionTarget inspectionTarget =
          resolver.resolveInspectionTarget(
              new StepId("inspect-table-summary"),
              workbook,
              tableCell,
              new WorkbookIntrospectionQuery.GetWorkbookSummary());
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
                  new TableSelector.ByName(new TableName("TextTable")),
                  "Item",
                  textCell("Missing")),
              "Amount");

      SemanticSelectorResolver.ResolvedInspectionTarget cellsResult =
          resolver.resolveInspectionTarget(
              new StepId("inspect-cells"),
              workbook,
              missingTarget,
              DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery());
      assertTrue(cellsResult.isShortCircuit());
      assertEquals(
          "Texts",
          assertInstanceOf(
                  SheetInspectionResult.CellsResult.class, cellsResult.shortCircuitResult())
              .sheetName()
              .value());

      SemanticSelectorResolver.ResolvedInspectionTarget hyperlinksResult =
          resolver.resolveInspectionTarget(
              new StepId("inspect-links"),
              workbook,
              missingTarget,
              new SheetIntrospectionQuery.GetHyperlinks());
      assertTrue(hyperlinksResult.isShortCircuit());
      assertEquals(
          "Texts",
          assertInstanceOf(
                  SheetInspectionResult.HyperlinksResult.class,
                  hyperlinksResult.shortCircuitResult())
              .sheetName()
              .value());

      SemanticSelectorResolver.ResolvedInspectionTarget commentsResult =
          resolver.resolveInspectionTarget(
              new StepId("inspect-comments"),
              workbook,
              missingTarget,
              new SheetIntrospectionQuery.GetComments());
      assertTrue(commentsResult.isShortCircuit());
      assertEquals(
          "Texts",
          assertInstanceOf(
                  SheetInspectionResult.CommentsResult.class, commentsResult.shortCircuitResult())
              .sheetName()
              .value());
    }
  }

  @Test
  void keyedResolutionTreatsNonMatchingSupportedScalarKeysAsZeroMatch() throws IOException {
    try (ExcelWorkbook workbook = workbookWithSemanticTables()) {
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-text-type-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("NumberTable")),
                          "Item",
                          textCell("Hosting")),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-blank-type-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("TextTable")),
                          "Item",
                          new CellValueSpec.Blank()),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-number-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("NumberTable")),
                          "Item",
                          new CellValueSpec.NumberValue(99.0d)),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-boolean-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("BooleanTable")),
                          "Item",
                          new CellValueSpec.BooleanValue(false)),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-formula-type-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("TextTable")),
                          "Item",
                          formulaCell("2+2")),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-formula-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("FormulaTable")),
                          "Item",
                          formulaCell("1+1")),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-number-type-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("TextTable")),
                          "Item",
                          new CellValueSpec.NumberValue(42.0d)),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
              .isShortCircuit());
      assertTrue(
          resolver
              .resolveInspectionTarget(
                  new StepId("inspect-boolean-type-miss"),
                  workbook,
                  new TableCellSelector.ByColumnName(
                      new TableRowSelector.ByKeyCell(
                          new TableSelector.ByName(new TableName("TextTable")),
                          "Item",
                          new CellValueSpec.BooleanValue(true)),
                      "Amount"),
                  DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery())
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
                              new TableSelector.ByName(new TableName("TextTable")),
                              "Item",
                              textCell("Missing")),
                          "Amount"),
                      new CellMutationAction.SetCell(textCell("Updated"))));
      assertTrue(missingRow.getMessage().contains("matched no rows"));

      IllegalArgumentException outOfRange =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveAssertionTarget(
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByIndex(
                              new TableSelector.ByName(new TableName("TextTable")), 4),
                          "Amount"),
                      new CellAssertion.DisplayValue("100")));
      assertTrue(outOfRange.getMessage().contains("outside the data-row bounds"));

      IllegalArgumentException missingTable =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveInspectionTarget(
                      new StepId("inspect-missing-table"),
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByName(new TableName("MissingTable")),
                              "Item",
                              textCell("Hosting")),
                          "Amount"),
                      DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery()));
      assertTrue(missingTable.getMessage().contains("table not found"));

      IllegalArgumentException missingTableOnExpectedSheet =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveInspectionTarget(
                      new StepId("inspect-missing-table-on-sheet"),
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByNameOnSheet(
                                  new TableName("TextTable"), new SheetName("Budget")),
                              "Item",
                              textCell("Hosting")),
                          "Amount"),
                      DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery()));
      assertTrue(
          missingTableOnExpectedSheet.getMessage().contains("table not found on expected sheet"));

      IllegalArgumentException missingNamedTableOnExpectedSheet =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  resolver.resolveInspectionTarget(
                      new StepId("inspect-missing-named-table-on-sheet"),
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByNameOnSheet(
                                  new TableName("MissingTable"), new SheetName("Budget")),
                              "Item",
                              textCell("Hosting")),
                          "Amount"),
                      DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery()));
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
                              new TableSelector.ByName(new TableName("TextTable")),
                              "MissingColumn",
                              textCell("Hosting")),
                          "Amount"),
                      new CellAssertion.DisplayValue("100")));
      assertTrue(missingColumn.getMessage().contains("does not contain column MissingColumn"));

      IllegalStateException unresolvedTextSource =
          assertThrows(
              IllegalStateException.class,
              () ->
                  resolver.resolveMutationTarget(
                      workbook,
                      new TableCellSelector.ByColumnName(
                          new TableRowSelector.ByKeyCell(
                              new TableSelector.ByName(new TableName("TextTable")),
                              "Item",
                              new CellValueSpec.Text(TextSourceSpec.utf8File("item.txt"))),
                          "Amount"),
                      new CellMutationAction.SetCell(textCell("Updated"))));
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
                workbook, target, new CellMutationAction.SetCell(textCell("Updated"))));
    assertEquals(expectedSheet, mutationTarget.sheetName().value());
    assertEquals(expectedAddress, mutationTarget.address().value());

    CellSelector.ByAddress assertionTarget =
        assertInstanceOf(
            CellSelector.ByAddress.class,
            resolver.resolveAssertionTarget(
                workbook, target, new CellAssertion.DisplayValue("100")));
    assertEquals(expectedSheet, assertionTarget.sheetName().value());
    assertEquals(expectedAddress, assertionTarget.address().value());

    CellStyleReport expectedStyle =
        InspectionResultCellReportSupport.toCellReport(
                workbook.sheet(expectedSheet).cells().snapshotCell(expectedAddress),
                DefaultGridGrindRequestExecutorTestSupport.projection(),
                false)
            .style()
            .orElseThrow();
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(
                    workbook,
                    target,
                    new CellMutationAction.SetHyperlink(
                        new HyperlinkTarget.Url("https://example.com"))))
            .address()
            .value());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(
                    workbook, target, new CellMutationAction.ClearHyperlink()))
            .address()
            .value());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(
                    workbook,
                    target,
                    new CellMutationAction.SetComment(
                        new CommentSpec(
                            TextSourceSpec.inline("Note"),
                            "Ada",
                            false,
                            java.util.Optional.empty(),
                            java.util.Optional.empty()))))
            .address()
            .value());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveMutationTarget(
                    workbook, target, new CellMutationAction.ClearComment()))
            .address()
            .value());

    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveAssertionTarget(
                    workbook,
                    target,
                    new CellAssertion.Value(
                        new dev.erst.gridgrind.core.model.value.CellValue.NumberValue(100.0d))))
            .address()
            .value());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveAssertionTarget(
                    workbook, target, new CellAssertion.FormulaText("2+2")))
            .address()
            .value());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver.resolveAssertionTarget(
                    workbook, target, new CellAssertion.CellStyle(expectedStyle)))
            .address()
            .value());

    SemanticSelectorResolver.ResolvedInspectionTarget inspectionTarget =
        resolver.resolveInspectionTarget(
            new StepId("inspect-cell"),
            workbook,
            target,
            DefaultGridGrindRequestExecutorTestSupport.allFacetCellsQuery());
    assertFalse(inspectionTarget.isShortCircuit());
    CellSelector.ByAddress inspectionAddress =
        assertInstanceOf(CellSelector.ByAddress.class, inspectionTarget.selector());
    assertEquals(expectedSheet, inspectionAddress.sheetName().value());
    assertEquals(expectedAddress, inspectionAddress.address().value());

    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver
                    .resolveInspectionTarget(
                        new StepId("inspect-links"),
                        workbook,
                        target,
                        new SheetIntrospectionQuery.GetHyperlinks())
                    .selector())
            .address()
            .value());
    assertEquals(
        expectedAddress,
        assertInstanceOf(
                CellSelector.ByAddress.class,
                resolver
                    .resolveInspectionTarget(
                        new StepId("inspect-comments"),
                        workbook,
                        target,
                        new SheetIntrospectionQuery.GetComments())
                    .selector())
            .address()
            .value());
  }

  private static TableCellSelector.ByColumnName textAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("TextTable")), "Item", textCell("Hosting")),
        "Amount");
  }

  private static TableCellSelector.ByColumnName textAmountCellTargetOnSheet() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByNameOnSheet(new TableName("TextTable"), new SheetName("Texts")),
            "Item",
            textCell("Hosting")),
        "Amount");
  }

  private static TableCellSelector.ByColumnName blankAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("BlankTable")),
            "Item",
            new CellValueSpec.Blank()),
        "Amount");
  }

  private static TableCellSelector.ByColumnName numberAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("NumberTable")),
            "Item",
            new CellValueSpec.NumberValue(42.0d)),
        "Amount");
  }

  private static TableCellSelector.ByColumnName booleanAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("BooleanTable")),
            "Item",
            new CellValueSpec.BooleanValue(true)),
        "Amount");
  }

  private static TableCellSelector.ByColumnName formulaAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("FormulaTable")), "Item", formulaCell("2+2")),
        "Amount");
  }

  private static TableCellSelector.ByColumnName wideAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("WideTable")), "Item", textCell("Hosting")),
        "Amount");
  }

  private static ExcelWorkbook workbookWithSemanticTables() throws IOException {
    ExcelWorkbook workbook = ExcelWorkbooks.create();
    addTable(
        workbook,
        "Texts",
        "TextTable",
        new CellValue.Text("Hosting"),
        new CellValue.NumberValue(100.0d));
    addTable(
        workbook, "Blanks", "BlankTable", new CellValue.Blank(), new CellValue.NumberValue(10.0d));
    addTable(
        workbook,
        "Numbers",
        "NumberTable",
        new CellValue.NumberValue(42.0d),
        new CellValue.NumberValue(200.0d));
    addTable(
        workbook,
        "Flags",
        "BooleanTable",
        new CellValue.BooleanValue(true),
        new CellValue.NumberValue(300.0d));
    addTable(
        workbook,
        "Formulas",
        "FormulaTable",
        new ResolvedFormula("2+2"),
        new CellValue.NumberValue(400.0d));
    addWideTable(workbook);
    workbook.getOrCreateSheet("Budget");
    workbook.sheet("Budget").cells().setCell("A1", new CellValue.Text("Owner"));
    return workbook;
  }

  private static void addTable(
      ExcelWorkbook workbook,
      String sheetName,
      String tableName,
      ResolvedCellWrite keyValue,
      ResolvedCellWrite amountValue) {
    workbook.getOrCreateSheet(sheetName);
    workbook.sheet(sheetName).cells().setCell("A1", new CellValue.Text("Item"));
    workbook.sheet(sheetName).cells().setCell("B1", new CellValue.Text("Amount"));
    workbook.sheet(sheetName).cells().setCell("A2", keyValue);
    workbook.sheet(sheetName).cells().setCell("B2", amountValue);
    workbook
        .tables()
        .setTable(
            new ExcelTableDefinition(
                tableName, sheetName, "A1:B2", false, new ExcelTableStyle.None()));
  }

  private static void addWideTable(ExcelWorkbook workbook) {
    workbook.getOrCreateSheet("Wide");
    workbook.sheet("Wide").cells().setCell("AA1", new CellValue.Text("Item"));
    workbook.sheet("Wide").cells().setCell("AB1", new CellValue.Text("Amount"));
    workbook.sheet("Wide").cells().setCell("AA2", new CellValue.Text("Hosting"));
    workbook.sheet("Wide").cells().setCell("AB2", new CellValue.NumberValue(500.0d));
    workbook
        .tables()
        .setTable(
            new ExcelTableDefinition(
                "WideTable", "Wide", "AA1:AB2", false, new ExcelTableStyle.None()));
  }
}
