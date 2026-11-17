package dev.erst.gridgrind.engine.drawing;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingCoordinateReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.problem.WorkbookContentLocation;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.util.Optional;
import java.util.function.Supplier;
import org.apache.xmlbeans.SimpleValue;
import org.apache.xmlbeans.XmlAnySimpleType;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.impl.values.XmlValueOutOfRangeException;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.main.STCoordinate;
import org.openxmlformats.schemas.drawingml.x2006.main.STCoordinateUnqualified;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STUniversalMeasure;

/** Reads native DrawingML coordinates without POI's narrowing or unit conversion. */
final class StoredDrawingAnchorReader {
  private StoredDrawingAnchorReader() {}

  static DrawingAnchorReport anchor(XmlObject shapeXml) {
    return read(
        "drawing.anchor",
        () ->
            switch (DrawingAdapter.parentAnchor(shapeXml).orElse(null)) {
              case CTTwoCellAnchor anchor ->
                  new DrawingAnchorReport.TwoCell(
                      marker(anchor.getFrom()),
                      marker(anchor.getTo()),
                      DrawingAdapter.behavior(anchor.getEditAs()));
              case CTOneCellAnchor anchor ->
                  new DrawingAnchorReport.OneCell(
                      marker(anchor.getFrom()),
                      extent(anchor.getExt(), true),
                      extent(anchor.getExt(), false),
                      ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
              case CTAbsoluteAnchor anchor -> {
                var position = anchor.getPos();
                if (position == null) throw invalid("drawing.anchor.pos", null);
                yield new DrawingAnchorReport.Absolute(
                    coordinate(position.xgetX(), "drawing.anchor.x"),
                    coordinate(position.xgetY(), "drawing.anchor.y"),
                    extent(anchor.getExt(), true),
                    extent(anchor.getExt(), false),
                    ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE);
              }
              case null -> throw invalid("drawing.anchor", null);
              default -> throw new IllegalStateException("Unsupported native drawing anchor");
            });
  }

  static DrawingMarkerReport marker(@Nullable CTMarker marker) {
    return read(
        "drawing.anchor.marker",
        () -> {
          if (marker == null) throw invalid("drawing.anchor.marker", null);
          return new DrawingMarkerReport(
              integer(marker.xgetCol(), "drawing.anchor.col"),
              integer(marker.xgetRow(), "drawing.anchor.row"),
              coordinate(marker.xgetColOff(), "drawing.anchor.colOff"),
              coordinate(marker.xgetRowOff(), "drawing.anchor.rowOff"));
        });
  }

  private static int integer(@Nullable XmlAnySimpleType scalar, String field) {
    return ((SimpleValue) requireValid(scalar, field)).getIntValue();
  }

  private static DrawingCoordinateReport coordinate(@Nullable STCoordinate scalar, String field) {
    var valid = requireValid(scalar, field);
    if (STCoordinateUnqualified.type.equals(valid.instanceType())) {
      return new DrawingCoordinateReport.Emu(((SimpleValue) valid).getLongValue());
    }
    if (STUniversalMeasure.type.equals(valid.instanceType())) {
      return new DrawingCoordinateReport.Measured(valid.getStringValue());
    }
    throw new IllegalStateException("Unsupported native coordinate member");
  }

  private static long extent(@Nullable CTPositiveSize2D size, boolean width) {
    String field = width ? "drawing.anchor.cx" : "drawing.anchor.cy";
    if (size == null) throw invalid(field, null);
    var scalar = width ? size.xgetCx() : size.xgetCy();
    return ((SimpleValue) requireValid(scalar, field)).getLongValue();
  }

  private static <T extends XmlAnySimpleType> T requireValid(@Nullable T scalar, String field) {
    if (scalar == null || !scalar.validate()) throw invalid(field, null);
    return scalar;
  }

  private static <T> T read(String field, Supplier<T> operation) {
    try {
      return operation.get();
    } catch (WorkbookContentRefusalException | UnexpectedNativeContentException exception) {
      throw exception;
    } catch (XmlValueOutOfRangeException exception) {
      throw invalid(field, exception);
    } catch (RuntimeException exception) {
      throw new UnexpectedNativeContentException(exception);
    }
  }

  private static WorkbookContentRefusalException invalid(String field, @Nullable Throwable cause) {
    return new WorkbookContentRefusalException(
        WorkbookContentIssue.INVALID_NATIVE_VALUE,
        Optional.of(
            new WorkbookContentLocation(
                Optional.empty(),
                Optional.empty(),
                Optional.of(field),
                Optional.empty(),
                Optional.empty())),
        cause);
  }
}
