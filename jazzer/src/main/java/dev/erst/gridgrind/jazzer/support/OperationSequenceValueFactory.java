package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelHeaderFooterText;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelPrintLayout;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import java.util.Objects;
import java.util.stream.Stream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Builds bounded protocol and engine value payloads shared across the Jazzer generators. */
final class OperationSequenceValueFactory {
  private static final String DRAWING_PICTURE_NAME = "OpsPicture";
  private static final String DRAWING_CHART_NAME = "OpsChart";
  private static final String PIVOT_TABLE_NAME = "OpsPivot";
  private static final String DRAWING_SHAPE_NAME = "OpsShape";
  private static final String DRAWING_CONNECTOR_NAME = "OpsConnector";
  private static final String DRAWING_EMBEDDED_OBJECT_NAME = "OpsEmbed";
  private static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private OperationSequenceValueFactory() {}

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
    return new PrintLayoutInput(
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

  static WorkflowStorage nextWorkflowStorage(
      String primarySheet, String secondarySheet, GridGrindFuzzData data) throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-jazzer-workflow-");
    Path sourcePath = directory.resolve("source.xlsx");
    Path saveAsPath = directory.resolve("output.xlsx");

    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 ->
          new WorkflowStorage(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.None(),
              directory);
      case 1 ->
          new WorkflowStorage(
              new WorkbookPlan.WorkbookSource.New(),
              new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString()),
              directory);
      case 2 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            directory);
      }
      case 3 -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
            directory);
      }
      default -> {
        writeExistingWorkbook(sourcePath, primarySheet, secondarySheet, data);
        yield new WorkflowStorage(
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

  static HyperlinkTarget nextHyperlinkTarget(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new HyperlinkTarget.Url("https://example.com/" + nextNamedRangeName(data, true));
      case 1 -> new HyperlinkTarget.Email(nextNamedRangeName(data, true) + "@example.com");
      case 2 -> new HyperlinkTarget.File("/tmp/" + nextNamedRangeName(data, true) + ".xlsx");
      default -> new HyperlinkTarget.Document("Budget!A1");
    };
  }

  static ExcelHyperlink nextExcelHyperlink(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelHyperlink.Url("https://example.com/" + nextNamedRangeName(data, true));
      case 1 -> new ExcelHyperlink.Email(nextNamedRangeName(data, true) + "@example.com");
      case 2 -> new ExcelHyperlink.File("/tmp/" + nextNamedRangeName(data, true) + ".xlsx");
      default -> new ExcelHyperlink.Document("Budget!A1");
    };
  }

  static CommentInput nextCommentInput(GridGrindFuzzData data) {
    return new CommentInput(
        TextSourceInput.inline("Note " + nextNamedRangeName(data, true)),
        "GridGrind",
        data.consumeBoolean());
  }

  static PictureInput nextPictureInput(GridGrindFuzzData data) {
    return new PictureInput(
        DRAWING_PICTURE_NAME,
        nextPictureDataInput(),
        nextDrawingAnchorInput(data),
        data.consumeBoolean() ? TextSourceInput.inline("Queue preview") : null);
  }

  static ChartInput nextChartInput(GridGrindFuzzData data) {
    return OperationSequenceChartFactory.nextChartInput(data);
  }

  static ShapeInput nextShapeInput(GridGrindFuzzData data) {
    if (data.consumeBoolean()) {
      return new ShapeInput(
          DRAWING_SHAPE_NAME,
          ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
          nextDrawingAnchorInput(data),
          data.consumeBoolean() ? "roundRect" : "rect",
          data.consumeBoolean() ? TextSourceInput.inline("Queue") : null);
    }
    return new ShapeInput(
        DRAWING_CONNECTOR_NAME,
        ExcelAuthoredDrawingShapeKind.CONNECTOR,
        nextDrawingAnchorInput(data),
        null,
        null);
  }

  static EmbeddedObjectInput nextEmbeddedObjectInput(GridGrindFuzzData data) {
    return new EmbeddedObjectInput(
        DRAWING_EMBEDDED_OBJECT_NAME,
        "Ops payload",
        "ops-payload.txt",
        "open",
        BinarySourceInput.inlineBase64(
            Base64.getEncoder()
                .encodeToString(
                    ("GridGrind payload " + data.consumeInt(0, 9))
                        .getBytes(StandardCharsets.UTF_8))),
        nextPictureDataInput(),
        nextDrawingAnchorInput(data));
  }

  static DrawingAnchorInput.TwoCell nextDrawingAnchorInput(GridGrindFuzzData data) {
    int firstColumn = data.consumeInt(0, 4);
    int firstRow = data.consumeInt(0, 8);
    int lastColumn = data.consumeInt(firstColumn + 1, firstColumn + 3);
    int lastRow = data.consumeInt(firstRow + 1, firstRow + 4);
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(firstColumn, firstRow, 0, 0),
        new DrawingMarkerInput(lastColumn, lastRow, 0, 0),
        nextDrawingAnchorBehavior(data));
  }

  static PictureDataInput nextPictureDataInput() {
    return new PictureDataInput(
        ExcelPictureFormat.PNG, BinarySourceInput.inlineBase64(PNG_PIXEL_BASE64));
  }

  static ExcelPictureDefinition nextExcelPictureDefinition(GridGrindFuzzData data) {
    return new ExcelPictureDefinition(
        DRAWING_PICTURE_NAME,
        new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        ExcelPictureFormat.PNG,
        nextExcelDrawingAnchor(data),
        data.consumeBoolean() ? "Queue preview" : null);
  }

  static ExcelChartDefinition nextExcelChartDefinition(GridGrindFuzzData data) {
    return OperationSequenceChartFactory.nextExcelChartDefinition(data);
  }

  static ExcelShapeDefinition nextExcelShapeDefinition(GridGrindFuzzData data) {
    if (data.consumeBoolean()) {
      return new ExcelShapeDefinition(
          DRAWING_SHAPE_NAME,
          ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
          nextExcelDrawingAnchor(data),
          data.consumeBoolean() ? "roundRect" : "rect",
          data.consumeBoolean() ? "Queue" : null);
    }
    return new ExcelShapeDefinition(
        DRAWING_CONNECTOR_NAME,
        ExcelAuthoredDrawingShapeKind.CONNECTOR,
        nextExcelDrawingAnchor(data),
        null,
        null);
  }

  static ExcelEmbeddedObjectDefinition nextExcelEmbeddedObjectDefinition(GridGrindFuzzData data) {
    return new ExcelEmbeddedObjectDefinition(
        DRAWING_EMBEDDED_OBJECT_NAME,
        "Ops payload",
        "ops-payload.txt",
        "open",
        new ExcelBinaryData(
            ("GridGrind payload " + data.consumeInt(0, 9)).getBytes(StandardCharsets.UTF_8)),
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        nextExcelDrawingAnchor(data));
  }

  static ExcelDrawingAnchor.TwoCell nextExcelDrawingAnchor(GridGrindFuzzData data) {
    DrawingAnchorInput.TwoCell anchor = nextDrawingAnchorInput(data);
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(
            anchor.from().columnIndex(),
            anchor.from().rowIndex(),
            anchor.from().dx(),
            anchor.from().dy()),
        new ExcelDrawingMarker(
            anchor.to().columnIndex(), anchor.to().rowIndex(), anchor.to().dx(), anchor.to().dy()),
        anchor.behavior());
  }

  static ExcelDrawingAnchorBehavior nextDrawingAnchorBehavior(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) & 0x3) {
      case 0 -> ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE;
      case 1 -> ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE;
      default -> ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE;
    };
  }

  static String nextDrawingObjectName(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 5) {
      case 0 -> DRAWING_PICTURE_NAME;
      case 1 -> DRAWING_CHART_NAME;
      case 2 -> DRAWING_SHAPE_NAME;
      case 3 -> DRAWING_CONNECTOR_NAME;
      default -> DRAWING_EMBEDDED_OBJECT_NAME;
    };
  }

  static String nextDrawingBinaryObjectName(GridGrindFuzzData data) {
    return data.consumeBoolean() ? DRAWING_PICTURE_NAME : DRAWING_EMBEDDED_OBJECT_NAME;
  }

  static DataValidationInput nextDataValidationInput(GridGrindFuzzData data) {
    return new DataValidationInput(
        data.consumeBoolean()
            ? new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done"))
            : new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new DataValidationPromptInput(
                TextSourceInput.inline("Status"),
                TextSourceInput.inline("Use an allowed value."),
                data.consumeBoolean())
            : null,
        data.consumeBoolean()
            ? new DataValidationErrorAlertInput(
                ExcelDataValidationErrorStyle.STOP,
                TextSourceInput.inline("Invalid"),
                TextSourceInput.inline("Use an allowed value."),
                data.consumeBoolean())
            : null);
  }

  static ExcelDataValidationDefinition nextExcelDataValidationDefinition(GridGrindFuzzData data) {
    return new ExcelDataValidationDefinition(
        data.consumeBoolean()
            ? new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done"))
            : new ExcelDataValidationRule.WholeNumber(
                ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
        data.consumeBoolean(),
        data.consumeBoolean(),
        data.consumeBoolean()
            ? new ExcelDataValidationPrompt(
                "Status", "Use an allowed value.", data.consumeBoolean())
            : null,
        data.consumeBoolean()
            ? new ExcelDataValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP,
                "Invalid",
                "Use an allowed value.",
                data.consumeBoolean())
            : null);
  }

  static ConditionalFormattingBlockInput nextConditionalFormattingInput(
      GridGrindFuzzData data, boolean validRange) {
    return new ConditionalFormattingBlockInput(
        List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        List.of(
            data.consumeBoolean()
                ? new ConditionalFormattingRuleInput.FormulaRule(
                    "A1>0", data.consumeBoolean(), nextDifferentialStyleInput(data))
                : new ConditionalFormattingRuleInput.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    null,
                    data.consumeBoolean(),
                    nextDifferentialStyleInput(data))));
  }

  static ExcelConditionalFormattingBlockDefinition nextExcelConditionalFormattingBlockDefinition(
      GridGrindFuzzData data, boolean validRange) {
    return new ExcelConditionalFormattingBlockDefinition(
        List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
        List.of(
            data.consumeBoolean()
                ? new ExcelConditionalFormattingRule.FormulaRule(
                    "A1>0", data.consumeBoolean(), nextExcelDifferentialStyle(data))
                : new ExcelConditionalFormattingRule.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    null,
                    data.consumeBoolean(),
                    nextExcelDifferentialStyle(data))));
  }

  static DifferentialStyleInput nextDifferentialStyleInput(GridGrindFuzzData data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    String fontColor = data.consumeBoolean() ? "#102030" : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    String fillColor = data.consumeBoolean() ? "#E0F0AA" : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new DifferentialStyleInput(
        numberFormat,
        bold,
        italic,
        null,
        java.util.Optional.ofNullable(fontColor),
        underline,
        strikeout,
        java.util.Optional.ofNullable(fillColor),
        java.util.Optional.empty());
  }

  static ExcelDifferentialStyle nextExcelDifferentialStyle(GridGrindFuzzData data) {
    boolean includeNumberFormat = data.consumeBoolean();
    Boolean bold = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean italic = data.consumeBoolean() ? Boolean.TRUE : null;
    String fontColor = data.consumeBoolean() ? "#102030" : null;
    Boolean underline = data.consumeBoolean() ? Boolean.TRUE : null;
    Boolean strikeout = data.consumeBoolean() ? Boolean.TRUE : null;
    String fillColor = data.consumeBoolean() ? "#E0F0AA" : null;
    String numberFormat =
        includeNumberFormat
                || Stream.of(bold, italic, fontColor, underline, strikeout, fillColor)
                    .allMatch(Objects::isNull)
            ? "0.00"
            : null;
    return new ExcelDifferentialStyle(
        numberFormat, bold, italic, null, fontColor, underline, strikeout, fillColor, null);
  }

  static RangeSelector nextRangeSelector(
      GridGrindFuzzData data, String sheetName, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new RangeSelector.AllOnSheet(sheetName);
      default ->
          new RangeSelector.ByRanges(
              sheetName,
              List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
    };
  }

  static ExcelRangeSelection nextExcelRangeSelection(GridGrindFuzzData data, boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelRangeSelection.All();
      default ->
          new ExcelRangeSelection.Selected(
              List.of(validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)));
    };
  }

  static String nextAutofilterRange(boolean validRange) {
    return validRange ? "E1:F3" : "BadRange";
  }

  static String nextCopySheetName(String sourceSheetName) {
    Objects.requireNonNull(sourceSheetName, "sourceSheetName must not be null");
    String base =
        sourceSheetName.length() <= 27 ? sourceSheetName : sourceSheetName.substring(0, 27);
    return base + "_B1";
  }

  static TableInput nextTableInput(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return new TableInput(
        tableName,
        sheetName,
        validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false),
        data.consumeBoolean(),
        nextTableStyleInput(data));
  }

  static ExcelTableDefinition nextExcelTableDefinition(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return new ExcelTableDefinition(
        tableName,
        sheetName,
        validRange ? "A1:B3" : FuzzDataDecoders.nextNonBlankRange(data, false),
        data.consumeBoolean(),
        nextExcelTableStyle(data));
  }

  static TableStyleInput nextTableStyleInput(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new TableStyleInput.None();
      default ->
          new TableStyleInput.Named(
              data.consumeBoolean() ? "TableStyleMedium2" : "TableStyleLight9",
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean());
    };
  }

  static ExcelTableStyle nextExcelTableStyle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new ExcelTableStyle.None();
      default ->
          new ExcelTableStyle.Named(
              data.consumeBoolean() ? "TableStyleMedium2" : "TableStyleLight9",
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean(),
              data.consumeBoolean());
    };
  }

  static TableSelector nextTableSelector(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new TableSelector.All();
      default ->
          new TableSelector.ByNames(
              List.of(
                  data.consumeBoolean()
                      ? nextTableName(data, true, primarySheet)
                      : nextTableName(data, true, secondarySheet)));
    };
  }

  static PivotTableSelector nextPivotTableSelector(
      GridGrindFuzzData data,
      String primarySheet,
      String secondarySheet,
      String pivotTableName,
      boolean validName) {
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> new PivotTableSelector.All();
      default ->
          data.consumeBoolean()
              ? new PivotTableSelector.ByNameOnSheet(
                  validName ? pivotTableName : nextPivotTableName(data, false), primarySheet)
              : new PivotTableSelector.ByNameOnSheet(
                  validName ? pivotTableName : nextPivotTableName(data, false), secondarySheet);
    };
  }

  static PivotTableInput nextPivotTableInput(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return new PivotTableInput(
        validName ? pivotTableName : nextPivotTableName(data, false),
        targetSheet,
        nextPivotTableSource(data, targetSheet, namedRangeName, tableName, validName, validRange),
        new PivotTableInput.Anchor(data.consumeBoolean() ? "F4" : "A6"),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new PivotTableInput.DataField(
                "Actual", ExcelPivotDataConsolidateFunction.SUM, "Total Actual", null)));
  }

  static PivotTableInput.Source nextPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new PivotTableInput.Source.Range(
              targetSheet, validRange ? "A1:C4" : FuzzDataDecoders.nextNonBlankRange(data, false));
      case 1 ->
          new PivotTableInput.Source.NamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false));
      default ->
          new PivotTableInput.Source.Table(
              validName ? tableName : nextTableName(data, false, targetSheet));
    };
  }

  static ExcelPivotTableDefinition nextExcelPivotTableDefinition(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return new ExcelPivotTableDefinition(
        validName ? pivotTableName : nextPivotTableName(data, false),
        targetSheet,
        nextExcelPivotTableSource(
            data, targetSheet, namedRangeName, tableName, validName, validRange),
        new ExcelPivotTableDefinition.Anchor(data.consumeBoolean() ? "F4" : "A6"),
        List.of("Month"),
        List.of(),
        List.of(),
        List.of(
            new ExcelPivotTableDefinition.DataField(
                "Actual", ExcelPivotDataConsolidateFunction.SUM, "Total Actual", null)));
  }

  static ExcelPivotTableDefinition.Source nextExcelPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 ->
          new ExcelPivotTableDefinition.Source.Range(
              targetSheet, validRange ? "A1:C4" : FuzzDataDecoders.nextNonBlankRange(data, false));
      case 1 ->
          new ExcelPivotTableDefinition.Source.NamedRange(
              validName ? namedRangeName : nextNamedRangeName(data, false));
      default ->
          new ExcelPivotTableDefinition.Source.Table(
              validName ? tableName : nextTableName(data, false, targetSheet));
    };
  }

  static String nextTableName(GridGrindFuzzData data, boolean valid, String sheetName) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (!valid) {
      return nextNamedRangeName(data, false);
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> sheetName + "Table";
      case 1 -> "BudgetTable";
      default -> "OpsTable";
    };
  }

  static ExcelComment nextExcelComment(GridGrindFuzzData data) {
    return new ExcelComment(
        "Note " + nextNamedRangeName(data, true), "GridGrind", data.consumeBoolean());
  }

  static String nextNamedRangeName(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");
    if (!valid) {
      return switch (selectorSlot(nextSelectorByte(data))) {
        case 0 -> "";
        case 1 -> "A1";
        case 2 -> "R1C1";
        case 3 -> "_xlnm.Print_Area";
        default -> "1Budget";
      };
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> "BudgetTotal";
      case 1 -> "LocalItem";
      case 2 -> "Report_Value";
      case 3 -> "Summary.Total";
      default -> "Name" + data.consumeInt(1, 9);
    };
  }

  static String nextPivotTableName(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");
    if (!valid) {
      return "";
    }
    return switch (selectorSlot(nextSelectorByte(data))) {
      case 0 -> PIVOT_TABLE_NAME;
      case 1 -> "Budget Pivot";
      default -> "Pivot " + data.consumeInt(1, 9);
    };
  }

  private static int nextSelectorByte(GridGrindFuzzData data) {
    return Byte.toUnsignedInt(data.consumeByte());
  }

  private static int selectorSlot(int selector) {
    return selector & 0x0F;
  }

  record WorkflowStorage(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      java.nio.file.Path cleanupRoot) {}
}
