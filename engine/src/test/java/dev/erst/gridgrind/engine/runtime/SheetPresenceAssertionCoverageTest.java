package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.io.IOException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage for SheetPresent/SheetAbsent assertion paths and related observation branches. */
class SheetPresenceAssertionCoverageTest {
  @Test
  void sheetPresentAndAbsentCoverAllSheetSelectorBranches() throws IOException {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    AssertionExecutor executor = new AssertionExecutor(workbookEngine, selectorResolver);
    WorkbookLocation location = new WorkbookLocation.UnsavedWorkbook();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("s1"),
                      new SheetSelector.ByName(new SheetName("Budget")),
                      new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      assertThrows(
          AssertionFailedException.class,
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("s2"),
                      new SheetSelector.ByName(new SheetName("Missing")),
                      new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("s3"),
                      new SheetSelector.ByName(new SheetName("Missing")),
                      new PresenceAssertion.SheetAbsent()),
                  workbook,
                  location));

      assertThrows(
          AssertionFailedException.class,
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("s4"),
                      new SheetSelector.ByName(new SheetName("Budget")),
                      new PresenceAssertion.SheetAbsent()),
                  workbook,
                  location));

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("s5"),
                      new SheetSelector.All(),
                      new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("s6"),
                      new SheetSelector.ByNames(List.of(new SheetName("Budget"))),
                      new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      workbook.getOrCreateSheet("Sales");
      assertThrows(
          AssertionFailedException.class,
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("s7"),
                      new SheetSelector.All(),
                      new PresenceAssertion.SheetAbsent()),
                  workbook,
                  location));
    }
  }

  @Test
  void namedRangeAbsentForMissingRangeCoversZeroMatchAndObservedCountPaths() throws IOException {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    AssertionExecutor executor = new AssertionExecutor(workbookEngine, selectorResolver);
    WorkbookLocation location = new WorkbookLocation.UnsavedWorkbook();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      new StepId("nr1"),
                      new NamedRangeSelector.ByName("MissingRange"),
                      new PresenceAssertion.NamedRangeAbsent()),
                  workbook,
                  location));
    }
  }
}
