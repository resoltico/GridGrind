package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Pivot-table authoring and source validation coverage. */
class ExcelPivotTableAuthoringCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  void setPivotTableRejectsInvalidAuthoringInputsAndReplacesSameSheetPivot() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      populatePivotSource(workbook, "Data");
      workbook.getOrCreateSheet("Report");
      workbook.getOrCreateSheet("OtherReport");

      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "C5",
              List.of(),
              List.of("Region"),
              List.of()));
      controller.setPivotTable(
          workbook,
          definition(
              "Sales Pivot",
              "Report",
              new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
              "E5",
              List.of("Owner"),
              List.of("Stage"),
              List.of()));

      ExcelPivotTableSnapshot.Supported replaced =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Supported.class,
              controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).getFirst());
      assertEquals(1, controller.pivotTables(workbook, new ExcelPivotTableSelection.All()).size());
      assertEquals("E5", replaced.anchor().topLeftAddress());
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
                          new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
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
                      new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
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
                      new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                      "C5",
                      List.of(),
                      List.of("Missing"),
                      List.of())));

      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "AmbiguousSource",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Data", "A1:D5")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "AmbiguousSource",
              new ExcelNamedRangeScope.SheetScope("Data"),
              new ExcelNamedRangeTarget("Data", "A1:D5")));

      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Named Range",
                      "Report",
                      new ExcelPivotTableDefinition.Source.NamedRange("MissingSource"),
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
                          new ExcelPivotTableDefinition.Source.NamedRange("AmbiguousSource"),
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
                      new ExcelPivotTableDefinition.Source.Table("MissingTable"),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));

      workbook
          .getOrCreateSheet("OneRow")
          .setRange(
              "A1:D1",
              List.of(
                  List.of(
                      ExcelCellValue.text("Region"),
                      ExcelCellValue.text("Stage"),
                      ExcelCellValue.text("Owner"),
                      ExcelCellValue.text("Amount"))));
      IllegalArgumentException oneRowFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "One Row",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("OneRow", "A1:D1"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(oneRowFailure.getMessage().contains("header row plus at least one data row"));

      ExcelSheet missingHeaderSheet = workbook.getOrCreateSheet("MissingHeader");
      missingHeaderSheet.setCell("A3", ExcelCellValue.text("North"));
      missingHeaderSheet.setCell("B3", ExcelCellValue.number(10));
      IllegalArgumentException missingHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Missing Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("MissingHeader", "A2:B3"),
                          "C5",
                          List.of(),
                          List.of("North"),
                          List.of())));
      assertTrue(missingHeaderFailure.getMessage().contains("missing its header row"));

      workbook
          .getOrCreateSheet("NumericHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.number(1), ExcelCellValue.text("Amount")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      IllegalArgumentException numericHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Numeric Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("NumericHeader", "A1:B2"),
                          "C5",
                          List.of(),
                          List.of("Amount"),
                          List.of())));
      assertFalse(numericHeaderFailure.getMessage().isBlank());

      workbook
          .getOrCreateSheet("BlankHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.text(" "), ExcelCellValue.text("Amount")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Blank Header",
                      "Report",
                      new ExcelPivotTableDefinition.Source.Range("BlankHeader", "A1:B2"),
                      "C5",
                      List.of(),
                      List.of("Amount"),
                      List.of())));

      ExcelSheet missingHeaderCellSheet = workbook.getOrCreateSheet("MissingHeaderCell");
      missingHeaderCellSheet.setCell("A1", ExcelCellValue.text("Region"));
      missingHeaderCellSheet.setCell("C1", ExcelCellValue.text("Amount"));
      missingHeaderCellSheet.setCell("A2", ExcelCellValue.text("North"));
      missingHeaderCellSheet.setCell("B2", ExcelCellValue.text("Plan"));
      missingHeaderCellSheet.setCell("C2", ExcelCellValue.number(10));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPivotTable(
                  workbook,
                  definition(
                      "Missing Header Cell",
                      "Report",
                      new ExcelPivotTableDefinition.Source.Range("MissingHeaderCell", "A1:C2"),
                      "C5",
                      List.of(),
                      List.of("Region"),
                      List.of())));

      workbook
          .getOrCreateSheet("DuplicateHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.text("Region"), ExcelCellValue.text("region")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      IllegalArgumentException duplicateHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Duplicate Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("DuplicateHeader", "A1:B2"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(duplicateHeaderFailure.getMessage().contains("unique case-insensitively"));

      workbook
          .getOrCreateSheet("BlankHeader")
          .setRange(
              "A1:B2",
              List.of(
                  List.of(ExcelCellValue.text(""), ExcelCellValue.text("Amount")),
                  List.of(ExcelCellValue.text("North"), ExcelCellValue.number(10))));
      IllegalArgumentException blankHeaderFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setPivotTable(
                      workbook,
                      definition(
                          "Blank Header",
                          "Report",
                          new ExcelPivotTableDefinition.Source.Range("BlankHeader", "A1:B2"),
                          "C5",
                          List.of(),
                          List.of("Region"),
                          List.of())));
      assertTrue(blankHeaderFailure.getMessage().contains("blank header cell"));
    }
  }
}
