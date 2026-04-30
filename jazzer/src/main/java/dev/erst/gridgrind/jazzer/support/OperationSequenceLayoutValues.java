package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.foundation.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Generates layout-oriented authored and engine values for operation-sequence fuzzing. */
final class OperationSequenceLayoutValues {
  private OperationSequenceLayoutValues() {}

  static PaneInput nextPaneInput(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new PaneInput.None();
      case 1 -> {
        int splitColumn = data.consumeInt(1, 3);
        yield new PaneInput.Frozen(
            splitColumn, 0, data.consumeInt(splitColumn, splitColumn + 2), 0);
      }
      case 2 -> {
        int splitRow = data.consumeInt(1, 3);
        yield new PaneInput.Frozen(0, splitRow, 0, data.consumeInt(splitRow, splitRow + 2));
      }
      case 3 -> {
        int splitColumn = data.consumeInt(1, 3);
        int splitRow = data.consumeInt(1, 3);
        yield new PaneInput.Frozen(
            splitColumn,
            splitRow,
            data.consumeInt(splitColumn, splitColumn + 2),
            data.consumeInt(splitRow, splitRow + 2));
      }
      default ->
          new PaneInput.Split(
              data.consumeInt(0, 2400),
              data.consumeInt(1, 2400),
              0,
              data.consumeInt(1, 4),
              nextProtocolPaneRegion(data));
    };
  }

  static ExcelSheetPane nextExcelSheetPane(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelSheetPane.None();
      case 1 -> {
        int splitColumn = data.consumeInt(1, 3);
        yield new ExcelSheetPane.Frozen(
            splitColumn, 0, data.consumeInt(splitColumn, splitColumn + 2), 0);
      }
      case 2 -> {
        int splitRow = data.consumeInt(1, 3);
        yield new ExcelSheetPane.Frozen(0, splitRow, 0, data.consumeInt(splitRow, splitRow + 2));
      }
      case 3 -> {
        int splitColumn = data.consumeInt(1, 3);
        int splitRow = data.consumeInt(1, 3);
        yield new ExcelSheetPane.Frozen(
            splitColumn,
            splitRow,
            data.consumeInt(splitColumn, splitColumn + 2),
            data.consumeInt(splitRow, splitRow + 2));
      }
      default ->
          new ExcelSheetPane.Split(
              data.consumeInt(0, 2400),
              data.consumeInt(1, 2400),
              0,
              data.consumeInt(1, 4),
              nextPaneRegion(data));
    };
  }

  static PrintLayoutInput nextPrintLayoutInput(GridGrindFuzzData data) {
    return PrintLayoutInput.withDefaultSetup(
        data.consumeBoolean() ? new PrintAreaInput.Range("A1:C20") : new PrintAreaInput.None(),
        data.consumeBoolean() ? ExcelPrintOrientation.LANDSCAPE : ExcelPrintOrientation.PORTRAIT,
        data.consumeBoolean()
            ? new PrintScalingInput.Fit(data.consumeInt(0, 2), data.consumeInt(0, 2))
            : new PrintScalingInput.Automatic(),
        data.consumeBoolean()
            ? new PrintTitleRowsInput.Band(0, data.consumeInt(0, 2))
            : new PrintTitleRowsInput.None(),
        data.consumeBoolean()
            ? new PrintTitleColumnsInput.Band(0, data.consumeInt(0, 2))
            : new PrintTitleColumnsInput.None(),
        new HeaderFooterTextInput(
            TextSourceInput.inline("L" + data.consumeInt(0, 9)),
            TextSourceInput.inline(""),
            TextSourceInput.inline("R" + data.consumeInt(0, 9))),
        new HeaderFooterTextInput(
            TextSourceInput.inline(""),
            TextSourceInput.inline("P" + data.consumeInt(0, 9)),
            TextSourceInput.inline("")));
  }

  static ExcelPrintLayout nextExcelPrintLayout(GridGrindFuzzData data) {
    return new ExcelPrintLayout(
        data.consumeBoolean()
            ? new ExcelPrintLayout.Area.Range("A1:C20")
            : new ExcelPrintLayout.Area.None(),
        data.consumeBoolean() ? ExcelPrintOrientation.LANDSCAPE : ExcelPrintOrientation.PORTRAIT,
        data.consumeBoolean()
            ? new ExcelPrintLayout.Scaling.Fit(data.consumeInt(0, 2), data.consumeInt(0, 2))
            : new ExcelPrintLayout.Scaling.Automatic(),
        data.consumeBoolean()
            ? new ExcelPrintLayout.TitleRows.Band(0, data.consumeInt(0, 2))
            : new ExcelPrintLayout.TitleRows.None(),
        data.consumeBoolean()
            ? new ExcelPrintLayout.TitleColumns.Band(0, data.consumeInt(0, 2))
            : new ExcelPrintLayout.TitleColumns.None(),
        new ExcelHeaderFooterText("L" + data.consumeInt(0, 9), "", "R" + data.consumeInt(0, 9)),
        new ExcelHeaderFooterText("", "P" + data.consumeInt(0, 9), ""));
  }

  static ExcelPaneRegion nextProtocolPaneRegion(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> ExcelPaneRegion.UPPER_LEFT;
      case 1 -> ExcelPaneRegion.UPPER_RIGHT;
      case 2 -> ExcelPaneRegion.LOWER_LEFT;
      default -> ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  static ExcelPaneRegion nextPaneRegion(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> ExcelPaneRegion.UPPER_LEFT;
      case 1 -> ExcelPaneRegion.UPPER_RIGHT;
      case 2 -> ExcelPaneRegion.LOWER_LEFT;
      default -> ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  static SheetCopyPosition nextSheetCopyPosition(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new SheetCopyPosition.AppendAtEnd();
      default -> new SheetCopyPosition.AtIndex(data.consumeInt(0, 2));
    };
  }

  static ExcelSheetCopyPosition nextExcelSheetCopyPosition(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelSheetCopyPosition.AppendAtEnd();
      default -> new ExcelSheetCopyPosition.AtIndex(data.consumeInt(0, 2));
    };
  }

  static List<String> nextSelectedSheetNames(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> List.of(primarySheet);
      case 1 -> List.of(secondarySheet);
      default -> List.of(primarySheet, secondarySheet);
    };
  }

  static ExcelSheetVisibility nextProtocolSheetVisibility(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> ExcelSheetVisibility.VISIBLE;
      case 1 -> ExcelSheetVisibility.HIDDEN;
      default -> ExcelSheetVisibility.VERY_HIDDEN;
    };
  }

  static ExcelSheetVisibility nextSheetVisibility(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> ExcelSheetVisibility.VISIBLE;
      case 1 -> ExcelSheetVisibility.HIDDEN;
      default -> ExcelSheetVisibility.VERY_HIDDEN;
    };
  }

  static SheetProtectionSettings nextProtocolSheetProtectionSettings(GridGrindFuzzData data) {
    return new SheetProtectionSettings(
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean());
  }

  static ExcelSheetProtectionSettings nextSheetProtectionSettings(GridGrindFuzzData data) {
    return new ExcelSheetProtectionSettings(
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean());
  }

  static OperationSequenceValueFactory.WorkflowStorage nextWorkflowStorage(
      String primarySheet, String secondarySheet, GridGrindFuzzData data) throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-jazzer-workflow-");
    Path sourcePath = directory.resolve("source.xlsx");
    Path saveAsPath = directory.resolve("output.xlsx");

    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 ->
          new OperationSequenceValueFactory.WorkflowStorage(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.None(),
              directory);
      case 1 ->
          new OperationSequenceValueFactory.WorkflowStorage(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString()),
              directory);
      case 2 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new OperationSequenceValueFactory.WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            directory);
      }
      case 3 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new OperationSequenceValueFactory.WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
            directory);
      }
      default -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new OperationSequenceValueFactory.WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString()),
            directory);
      }
    };
  }

  static void writeExistingWorkbook(
      Path sourcePath, String primarySheet, String secondarySheet, GridGrindFuzzData data)
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var primary = workbook.createSheet(primarySheet);
      primary.createRow(0).createCell(0).setCellValue("Month");
      primary.getRow(0).createCell(1).setCellValue("Plan");
      primary.getRow(0).createCell(2).setCellValue("Actual");
      primary.createRow(1).createCell(0).setCellValue("Jan");
      primary.getRow(1).createCell(1).setCellValue(2.0d);
      primary.getRow(1).createCell(2).setCellValue(4.0d);
      primary.getRow(1).createCell(3).setCellFormula("B2*2");
      primary.createRow(2).createCell(0).setCellValue("Feb");
      primary.getRow(2).createCell(1).setCellValue(3.0d);
      primary.getRow(2).createCell(2).setCellValue(6.0d);
      primary.createRow(3).createCell(0).setCellValue("Mar");
      primary.getRow(3).createCell(1).setCellValue(5.0d);
      primary.getRow(3).createCell(2).setCellValue(7.0d);
      primary.createRow(4).createCell(4).setCellValue("Queue");
      primary.getRow(4).createCell(5).setCellValue("Owner");
      primary.createRow(5).createCell(4).setCellValue("seed");
      primary.getRow(5).createCell(5).setCellValue("GridGrind");
      if (data.consumeBoolean()) {
        var secondary = workbook.createSheet(secondarySheet);
        secondary.createRow(0).createCell(0).setCellValue("Month");
        secondary.getRow(0).createCell(1).setCellValue("Plan");
        secondary.getRow(0).createCell(2).setCellValue("Actual");
        secondary.createRow(1).createCell(0).setCellValue("Jan");
        secondary.getRow(1).createCell(1).setCellValue(3.0d);
        secondary.getRow(1).createCell(2).setCellValue(9.0d);
        secondary.getRow(1).createCell(3).setCellFormula("B2*3");
        secondary.createRow(2).createCell(0).setCellValue("Feb");
        secondary.getRow(2).createCell(1).setCellValue(4.0d);
        secondary.getRow(2).createCell(2).setCellValue(8.0d);
        secondary.createRow(3).createCell(0).setCellValue("Mar");
        secondary.getRow(3).createCell(1).setCellValue(6.0d);
        secondary.getRow(3).createCell(2).setCellValue(10.0d);
      }
      Files.createDirectories(sourcePath.getParent());
      try (OutputStream outputStream = Files.newOutputStream(sourcePath)) {
        workbook.write(outputStream);
      }
    }
  }
}
