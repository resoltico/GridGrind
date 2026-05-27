package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
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
                      "s1",
                      new SheetSelector.ByName("Budget"),
                      new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      assertThrows(
          AssertionFailedException.class,
          () ->
              executor.execute(
                  new AssertionStep(
                      "s2",
                      new SheetSelector.ByName("Missing"),
                      new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      "s3",
                      new SheetSelector.ByName("Missing"),
                      new PresenceAssertion.SheetAbsent()),
                  workbook,
                  location));

      assertThrows(
          AssertionFailedException.class,
          () ->
              executor.execute(
                  new AssertionStep(
                      "s4",
                      new SheetSelector.ByName("Budget"),
                      new PresenceAssertion.SheetAbsent()),
                  workbook,
                  location));

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      "s5", new SheetSelector.All(), new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      assertDoesNotThrow(
          () ->
              executor.execute(
                  new AssertionStep(
                      "s6",
                      new SheetSelector.ByNames(List.of("Budget")),
                      new PresenceAssertion.SheetPresent()),
                  workbook,
                  location));

      workbook.getOrCreateSheet("Sales");
      assertThrows(
          AssertionFailedException.class,
          () ->
              executor.execute(
                  new AssertionStep(
                      "s7", new SheetSelector.All(), new PresenceAssertion.SheetAbsent()),
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
                      "nr1",
                      new NamedRangeSelector.ByName("MissingRange"),
                      new PresenceAssertion.NamedRangeAbsent()),
                  workbook,
                  location));
    }
  }
}
