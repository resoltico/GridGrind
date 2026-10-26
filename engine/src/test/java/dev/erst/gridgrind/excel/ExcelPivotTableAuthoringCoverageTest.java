package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Pivot-table authoring and source validation coverage. */
class ExcelPivotTableAuthoringCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  void setPivotTableRejectsInvalidAuthoringInputsAndReplacesSameSheetPivot() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      workbook.getOrCreateSheet("OtherReport");

      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot",
              "Report",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "C5",
              List.of(),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot",
              "Report",
              new PivotTableSpec.PivotSourceSpec.Range(
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
              "E5",
              List.of("Owner"),
              List.of("Stage"),
              List.of()));

      PivotTableReport.Supported replaced =
          assertInstanceOf(
              PivotTableReport.Supported.class,
              controller.pivotTables(workbook, new PivotTableSelector.All()).getFirst());
      assertEquals(1, controller.pivotTables(workbook, new PivotTableSelector.All()).size());
      assertEquals("E5", replaced.anchor().topLeftAddress().value());
      assertEquals(List.of("Stage"), fieldNames(replaced.rowLabels()));
      assertEquals(List.of("Owner"), fieldNames(replaced.reportFilters()));

      IllegalArgumentException differentSheetFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Sales Pivot",
                          "OtherReport",
                          new PivotTableSpec.PivotSourceSpec.Range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(differentSheetFailure.getMessage().contains("different sheet"));

      assertThrows(
          SheetNotFoundException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Report",
                      "MissingReport",
                      new PivotTableSpec.PivotSourceSpec.Range(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Column",
                      "Report",
                      new PivotTableSpec.PivotSourceSpec.Range(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                      "C5",
                      List.of(),
                      List.of("Missing"),
                      List.of())));

      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("AmbiguousSource"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5"))));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("AmbiguousSource"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data")),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5"))));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Named Range",
                      "Report",
                      new PivotTableSpec.PivotSourceSpec.NamedRange(
                          new dev.erst.gridgrind.core.vocabulary.DefinedName("MissingSource")),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));
      IllegalArgumentException ambiguousNamedRangeFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Ambiguous Named Range",
                          "Report",
                          new PivotTableSpec.PivotSourceSpec.NamedRange(
                              new dev.erst.gridgrind.core.vocabulary.DefinedName(
                                  "AmbiguousSource")),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(ambiguousNamedRangeFailure.getMessage().contains("ambiguous"));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Table",
                      "Report",
                      new PivotTableSpec.PivotSourceSpec.Table(
                          new dev.erst.gridgrind.core.vocabulary.TableName("MissingTable")),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));

      workbook
          .getOrCreateSheet("OneRow")
          .cells()
          .setRange(
              "A1:D1",
              List.of(
                  List.of(
                      new CellValue.Text("Region"),
                      new CellValue.Text("Stage"),
                      new CellValue.Text("Owner"),
                      new CellValue.Text("Amount"))));
      IllegalArgumentException oneRowFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "One Row",
                          "Report",
                          new PivotTableSpec.PivotSourceSpec.Range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("OneRow"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D1")),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(oneRowFailure.getMessage().contains("header row plus at least one data row"));

      ExcelSheet missingHeaderSheet = workbook.getOrCreateSheet("MissingHeader");
      missingHeaderSheet.cells().setCell("A3", new CellValue.Text("North"));
      missingHeaderSheet.cells().setCell("B3", new CellValue.NumberValue(10));
      IllegalArgumentException missingHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Missing Header",
                          "Report",
                          new PivotTableSpec.PivotSourceSpec.Range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("MissingHeader"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:B3")),
                          "C5",
                          List.of(),
                          List.of("North"),
                          List.of())));
      assertTrue(missingHeaderFailure.getMessage().contains("missing its header row"));

      workbook
          .getOrCreateSheet("NumericHeader")
          .cells()
          .setRange(
              "A1:B2",
              List.of(
                  List.of(new CellValue.NumberValue(1), new CellValue.Text("Amount")),
                  List.of(new CellValue.Text("North"), new CellValue.NumberValue(10))));
      IllegalArgumentException numericHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Numeric Header",
                          "Report",
                          new PivotTableSpec.PivotSourceSpec.Range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("NumericHeader"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2")),
                          "C5",
                          List.of(),
                          List.of("Amount"),
                          List.of())));
      assertFalse(numericHeaderFailure.getMessage().isBlank());

      workbook
          .getOrCreateSheet("BlankHeader")
          .cells()
          .setRange(
              "A1:B2",
              List.of(
                  List.of(new CellValue.Text(" "), new CellValue.Text("Amount")),
                  List.of(new CellValue.Text("North"), new CellValue.NumberValue(10))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Blank Header",
                      "Report",
                      new PivotTableSpec.PivotSourceSpec.Range(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("BlankHeader"),
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2")),
                      "C5",
                      List.of(),
                      List.of("Amount"),
                      List.of())));

      ExcelSheet missingHeaderCellSheet = workbook.getOrCreateSheet("MissingHeaderCell");
      missingHeaderCellSheet.cells().setCell("A1", new CellValue.Text("Region"));
      missingHeaderCellSheet.cells().setCell("C1", new CellValue.Text("Amount"));
      missingHeaderCellSheet.cells().setCell("A2", new CellValue.Text("North"));
      missingHeaderCellSheet.cells().setCell("B2", new CellValue.Text("Plan"));
      missingHeaderCellSheet.cells().setCell("C2", new CellValue.NumberValue(10));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Header Cell",
                      "Report",
                      new PivotTableSpec.PivotSourceSpec.Range(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("MissingHeaderCell"),
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:C2")),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));

      workbook
          .getOrCreateSheet("DuplicateHeader")
          .cells()
          .setRange(
              "A1:B2",
              List.of(
                  List.of(new CellValue.Text("Region"), new CellValue.Text("region")),
                  List.of(new CellValue.Text("North"), new CellValue.NumberValue(10))));
      IllegalArgumentException duplicateHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Duplicate Header",
                          "Report",
                          new PivotTableSpec.PivotSourceSpec.Range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("DuplicateHeader"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2")),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(duplicateHeaderFailure.getMessage().contains("unique case-insensitively"));

      workbook
          .getOrCreateSheet("BlankHeader")
          .cells()
          .setRange(
              "A1:B2",
              List.of(
                  List.of(new CellValue.Text(""), new CellValue.Text("Amount")),
                  List.of(new CellValue.Text("North"), new CellValue.NumberValue(10))));
      IllegalArgumentException blankHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Blank Header",
                          "Report",
                          new PivotTableSpec.PivotSourceSpec.Range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("BlankHeader"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2")),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(blankHeaderFailure.getMessage().contains("blank header cell"));
    }
  }
}
