package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.source.*;
import dev.erst.gridgrind.core.vocabulary.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelShapeDefinition;
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

  static CommentSpec nextCommentInput(GridGrindFuzzData data) {
    return CommentSpec.plain(
        TextSourceSpec.inline("Note " + nextNamedRangeName(data, true)),
        "GridGrind",
        data.consumeBoolean());
  }

  static PictureSpec nextPictureInput(GridGrindFuzzData data) {
    return new PictureSpec(
        DRAWING_PICTURE_NAME,
        nextPictureDataInput(),
        nextDrawingAnchorInput(data),
        data.consumeBoolean()
            ? Optional.of(TextSourceSpec.inline("Queue preview"))
            : Optional.empty());
  }

  static ChartSpec nextChartInput(GridGrindFuzzData data) {
    return OperationSequenceChartFactory.nextChartInput(data);
  }

  static ShapeSpec nextShapeInput(GridGrindFuzzData data) {
    if (data.consumeBoolean()) {
      return new ShapeSpec.SimpleShape(
          DRAWING_SHAPE_NAME,
          nextDrawingAnchorInput(data),
          data.consumeBoolean() ? "roundRect" : "rect",
          data.consumeBoolean() ? Optional.of(TextSourceSpec.inline("Queue")) : Optional.empty());
    }
    return new ShapeSpec.Connector(DRAWING_CONNECTOR_NAME, nextDrawingAnchorInput(data));
  }

  static EmbeddedObjectSpec nextEmbeddedObjectInput(GridGrindFuzzData data) {
    return new EmbeddedObjectSpec(
        DRAWING_EMBEDDED_OBJECT_NAME,
        "Ops payload",
        "ops-payload.txt",
        "open",
        BinarySourceSpec.inlineBase64(
            Base64.getEncoder()
                .encodeToString(
                    ("GridGrind payload " + data.consumeInt(0, 9))
                        .getBytes(StandardCharsets.UTF_8))),
        nextPictureDataInput(),
        nextDrawingAnchorInput(data));
  }

  static DrawingAnchorSpec.TwoCell nextDrawingAnchorInput(GridGrindFuzzData data) {
    int firstColumn = data.consumeInt(0, 4);
    int firstRow = data.consumeInt(0, 8);
    int lastColumn = data.consumeInt(firstColumn + 1, firstColumn + 3);
    int lastRow = data.consumeInt(firstRow + 1, firstRow + 4);
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(firstColumn, firstRow, 0, 0),
        new DrawingMarkerSpec(lastColumn, lastRow, 0, 0),
        nextDrawingAnchorBehavior(data));
  }

  static PictureDataSpec nextPictureDataInput() {
    return new PictureDataSpec(
        ExcelPictureFormat.PNG, BinarySourceSpec.inlineBase64(PNG_PIXEL_BASE64));
  }

  static ExcelPictureDefinition nextExcelPictureDefinition(GridGrindFuzzData data) {
    return new ExcelPictureDefinition(
        DRAWING_PICTURE_NAME,
        new ExcelBinaryData(Base64.getDecoder().decode(PNG_PIXEL_BASE64)),
        ExcelPictureFormat.PNG,
        nextDrawingAnchorInput(data),
        data.consumeBoolean() ? Optional.of("Queue preview") : Optional.empty());
  }

  static ResolvedChart nextResolvedChart(GridGrindFuzzData data) {
    return OperationSequenceChartFactory.nextResolvedChart(data);
  }

  static ExcelShapeDefinition nextExcelShapeDefinition(GridGrindFuzzData data) {
    if (data.consumeBoolean()) {
      return new ExcelShapeDefinition.SimpleShape(
          DRAWING_SHAPE_NAME,
          nextDrawingAnchorInput(data),
          data.consumeBoolean() ? "roundRect" : "rect",
          data.consumeBoolean() ? Optional.of("Queue") : Optional.empty());
    }
    return new ExcelShapeDefinition.Connector(DRAWING_CONNECTOR_NAME, nextDrawingAnchorInput(data));
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
        nextDrawingAnchorInput(data));
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

  static ResolvedComment nextExcelComment(GridGrindFuzzData data) {
    return new ResolvedComment(
        "Note " + nextNamedRangeName(data, true), "GridGrind", data.consumeBoolean());
  }
}
