package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import java.util.List;

/** Shared DSL helpers for the contract-owned generated example registry. */
final class ExamplePlanSupport {
  private ExamplePlanSupport() {}

  static GridGrindShippedExamples.ShippedExample example(
      String id, String fileName, String summary, WorkbookPlan plan) {
    return new GridGrindShippedExamples.ShippedExample(id, fileName, summary, plan);
  }

  static WorkbookPlan plan(
      String planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      WorkbookStep... steps) {
    return new WorkbookPlan(
        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
        planId,
        source,
        persistence,
        execution,
        null,
        List.of(steps));
  }

  static WorkbookPlan.WorkbookPersistence.SaveAs saveAs(String path) {
    return new WorkbookPlan.WorkbookPersistence.SaveAs(path);
  }

  static WorkbookSelector.Current workbook() {
    return new WorkbookSelector.Current();
  }

  static SheetSelector.ByName sheet(String name) {
    return new SheetSelector.ByName(name);
  }

  static SheetSelector.ByNames sheets(String... names) {
    return new SheetSelector.ByNames(List.of(names));
  }

  static CellSelector.ByAddress cell(String sheetName, String address) {
    return new CellSelector.ByAddress(sheetName, address);
  }

  static CellSelector.ByAddresses cells(String sheetName, String... addresses) {
    return new CellSelector.ByAddresses(sheetName, List.of(addresses));
  }

  static RangeSelector.ByRange range(String sheetName, String range) {
    return new RangeSelector.ByRange(sheetName, range);
  }

  static RangeSelector.RectangularWindow window(
      String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new RangeSelector.RectangularWindow(sheetName, topLeftAddress, rowCount, columnCount);
  }

  static DrawingAnchorInput.TwoCell anchor(int fromColumn, int fromRow, int toColumn, int toRow) {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(fromColumn, fromRow),
        new DrawingMarkerInput(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  static MutationStep step(String stepId, Selector target, MutationAction action) {
    return new MutationStep(stepId, target, action);
  }

  static InspectionStep read(String stepId, Selector target, InspectionQuery query) {
    return new InspectionStep(stepId, target, query);
  }

  static AssertionStep assertStep(String stepId, Selector target, Assertion assertion) {
    return new AssertionStep(stepId, target, assertion);
  }

  @SafeVarargs
  static List<List<CellInput>> rows(List<CellInput>... rows) {
    return List.of(rows);
  }

  static List<CellInput> row(CellInput... cells) {
    return List.of(cells);
  }

  static CellInput.Text text(String value) {
    return new CellInput.Text(TextSourceInput.inline(value));
  }

  static CellInput.Formula formula(String value) {
    return new CellInput.Formula(TextSourceInput.inline(value));
  }

  static CellInput.Numeric number(double value) {
    return new CellInput.Numeric(value);
  }

  static CellInput.BooleanValue bool(boolean value) {
    return new CellInput.BooleanValue(value);
  }
}
