package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.ResolvedCellText;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ResolvedInspectionStep;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.problem.FieldValidationAddressRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Observed table labels stay factual while executable table geometry follows the core parser. */
class SemanticTableGeometryTest {
  @ParameterizedTest
  @CsvSource({"B2:A1,RANGE_ORDER", "A1:B2:C3,RANGE_RECTANGULAR_SYNTAX"})
  void malformedObservedTableRangesCannotResolveTableKeys(
      String rawRange, FieldValidationAddressRule expectedRule) throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      var table = populateTable(workbook);
      table.getCTTable().setRef(rawRange);
      var engine = new WorkbookExecutionEngine();
      assertEquals(rawRange, observedRange(engine, workbook));
      var failure =
          assertThrows(
              InvalidRequestException.class,
              () ->
                  new SemanticSelectorResolver(engine)
                      .resolveAssertionTarget(
                          workbook,
                          keyTarget(),
                          new CellAssertion.Value(new CellValue.NumberValue(42))));
      assertEquals(
          expectedRule,
          assertInstanceOf(FieldValidationProblem.class, failure.requestProblem()).rule());
    }
  }

  @Test
  void anchoredObservedTableRangeResolvesCanonicalCoordinatesWithoutRewritingTheLabel()
      throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      var table = populateTable(workbook);
      String rawRange = "$a$1:$b$2";
      table.getCTTable().setRef(rawRange);
      var engine = new WorkbookExecutionEngine();
      assertEquals(rawRange, observedRange(engine, workbook));
      var resolved =
          assertInstanceOf(
              CellSelector.ByAddress.class,
              new SemanticSelectorResolver(engine)
                  .resolveAssertionTarget(
                      workbook,
                      keyTarget(),
                      new CellAssertion.Value(new CellValue.NumberValue(42))));
      assertEquals(new SheetName("Grid"), resolved.sheetName());
      assertEquals(new CellAddress(1, 1), resolved.address());
      assertEquals(rawRange, observedRange(engine, workbook));
    }
  }

  private static XSSFTable populateTable(XssfWorkbook workbook) {
    var sheet = workbook.getOrCreateSheet("Grid");
    sheet
        .cells()
        .setRange(
            new RangeRef(0, 1, 0, 1),
            List.of(
                List.<ResolvedCellWrite>of(
                    new ResolvedCellText("Item"), new ResolvedCellText("Amount")),
                List.<ResolvedCellWrite>of(
                    new ResolvedCellText("One"), new CellValueSpec.NumberValue(42))));
    new TableAdapter()
        .setTable(
            workbook,
            new ResolvedTable(
                new TableName("ObservedTable"),
                new SheetName("Grid"),
                new RangeRef(0, 1, 0, 1),
                false,
                new TableStyleSpec.None()));
    return workbook.xssfWorkbook().getSheet("Grid").getTables().getFirst();
  }

  private static TableCellSelector.ByColumnName keyTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("ObservedTable")),
            "Item",
            new CellValueSpec.Text(TextSourceSpec.inline("One"))),
        "Amount");
  }

  private static String observedRange(WorkbookExecutionEngine engine, XssfWorkbook workbook) {
    var result =
        assertInstanceOf(
            WorkbookAssetInspectionResult.TablesResult.class,
            engine
                .read(
                    workbook,
                    new ResolvedInspectionStep(
                        new StepId("observe-table"),
                        new TableSelector.ByName(new TableName("ObservedTable")),
                        new WorkbookAssetIntrospectionQuery.GetTables()))
                .getFirst());
    return result.tables().getFirst().range().value();
  }
}
