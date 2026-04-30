package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.foundation.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;

/** Generates authored and engine drawing values for operation-sequence fuzzing. */
final class OperationSequenceDrawingValues {
  private OperationSequenceDrawingValues() {}

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
    return CommentInput.plain(
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

  static ExcelComment nextExcelComment(GridGrindFuzzData data) {
    return new ExcelComment(
        "Note " + nextNamedRangeName(data, true), "GridGrind", data.consumeBoolean());
  }
}
