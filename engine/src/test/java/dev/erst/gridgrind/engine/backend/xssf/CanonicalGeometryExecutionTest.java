package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.ResolvedCellText;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedInspectionStep;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Exercises canonical coordinates across mutation dispatch, native storage and ordered reads. */
class CanonicalGeometryExecutionTest {
  @Test
  void mutationsAndOrderedReadsUseTheSameCanonicalCoordinates() throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      var name = new SheetName("Grid");
      var sheet = workbook.getOrCreateSheet(name.value());
      var first = new CellAddress(3, 25);
      var last = new CellAddress(4, 26);
      var single = new CellAddress(7, 27);
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(name, new RangeRef(3, 4, 25, 26)),
              new ResolvedMutationAction.SetRange(
                  List.of(
                      List.<ResolvedCellWrite>of(
                          new ResolvedCellText("one"), new ResolvedCellText("two")),
                      List.<ResolvedCellWrite>of(
                          new ResolvedCellText("three"), new ResolvedCellText("four"))))));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(name, single),
              new ResolvedMutationAction.SetCell(new ResolvedCellText("end"))));
      assertEquals("one", sheet.xssfSheet().getRow(3).getCell(25).getStringCellValue());
      assertEquals("four", sheet.xssfSheet().getRow(4).getCell(26).getStringCellValue());
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(name, new RangeRef(4, 4, 26, 26)),
              new CellMutationAction.ClearRange()));
      var read =
          new WorkbookExecutionEngine()
              .read(
                  workbook,
                  new ResolvedInspectionStep(
                      new StepId("ordered"),
                      new CellSelector.ByAddresses(name, List.of(single, first, last)),
                      new SheetIntrospectionQuery.GetCells()));
      var cells =
          assertInstanceOf(SheetInspectionResult.CellsResult.class, read.getFirst()).cells();
      assertEquals(List.of(single, first, last), cells.stream().map(CellReport::address).toList());
      assertEquals(new CellValue.Text("end"), cells.get(0).value().orElseThrow());
      assertEquals(new CellValue.Text("one"), cells.get(1).value().orElseThrow());
      assertInstanceOf(CellReport.BlankReport.class, cells.get(2));
      assertEquals("end", sheet.cells().text(new CellAddress("$ab$8")));
    }
  }

  @Test
  void typedArrayAndWindowBoundariesRetainIndicesWithoutOverflow() throws Exception {
    try (var workbook = XssfWorkbooks.create()) {
      var cells = workbook.getOrCreateSheet("Grid").cells();
      var range = new RangeRef(10, 11, 1, 1);
      cells.setArrayFormula(range, new ResolvedArrayFormula("ROW(B11:B12)"));
      assertEquals("B11:B12", cells.arrayFormulas().getFirst().range().value());
      cells.clearArrayFormula(new CellAddress(11, 1));
      assertEquals(List.of(), cells.arrayFormulas());
      assertThrows(
          IllegalArgumentException.class,
          () -> cells.window(new CellAddress(1, 1), Integer.MAX_VALUE, 1));
      assertThrows(
          IllegalArgumentException.class,
          () -> cells.window(new CellAddress(1, 1), 1, Integer.MAX_VALUE));
    }
  }
}
