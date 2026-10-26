package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchorSupport;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;

/** Distinguishes authored rectangle validation from factual drawing coordinates. */
class DrawingGeometryContractTest {
  @Test
  void observedReversedCoordinatesAreReportedWhileAuthoredRectanglesRejectThem() {
    CTDrawing drawing = CTDrawing.Factory.newInstance();
    var anchor = drawing.addNewTwoCellAnchor();
    var from = anchor.addNewFrom();
    from.setCol(4);
    from.setRow(5);
    from.setColOff(3);
    from.setRowOff(7);
    var to = anchor.addNewTo();
    to.setCol(1);
    to.setRow(2);
    to.setColOff(0);
    to.setRowOff(0);
    var shape = anchor.addNewSp();

    var report =
        assertInstanceOf(
            DrawingAnchorReport.TwoCell.class, ExcelDrawingAnchorSupport.snapshotAnchor(shape));
    assertEquals(4, report.from().columnIndex());
    assertEquals(5, report.from().rowIndex());
    assertEquals(3, report.from().dx());
    assertEquals(7, report.from().dy());
    assertEquals(1, report.to().columnIndex());
    assertEquals(2, report.to().rowIndex());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DrawingAnchorSpec.TwoCell(
                new DrawingMarkerSpec(4, 5, 3, 7),
                new DrawingMarkerSpec(1, 2, 0, 0),
                ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE));
  }
}
