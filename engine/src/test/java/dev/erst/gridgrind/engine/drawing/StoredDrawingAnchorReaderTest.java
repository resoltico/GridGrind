package dev.erst.gridgrind.engine.drawing;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingCoordinateReport;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.io.IOException;
import java.lang.reflect.Proxy;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;

/** Native simple values are validated before narrowing, with VML offsets retaining pixel units. */
class StoredDrawingAnchorReaderTest {
  @Test
  void drawingMlKeepsLargeSignedAndMeasuredCoordinatesWithoutMutation() {
    var drawing = CTDrawing.Factory.newInstance();
    var anchor = drawing.addNewTwoCellAnchor();
    var from = anchor.addNewFrom();
    from.setCol(Integer.MAX_VALUE);
    from.setRow(Integer.MAX_VALUE);
    from.setColOff(2147483648L);
    from.setRowOff("-01.00cm");
    var to = anchor.addNewTo();
    to.setCol(0);
    to.setRow(0);
    to.setColOff(-1L);
    to.setRowOff(17L);
    var shape = anchor.addNewSp();
    assertTrue(from.validate());
    assertTrue(to.validate());
    String before = drawing.xmlText();
    var report =
        assertInstanceOf(DrawingAnchorReport.TwoCell.class, DrawingAdapter.snapshotAnchor(shape));
    assertEquals(new DrawingCoordinateReport.Emu(2147483648L), report.from().dx());
    assertEquals(new DrawingCoordinateReport.Measured("-01.00cm"), report.from().dy());
    assertEquals(new DrawingCoordinateReport.Emu(-1), report.to().dx());
    assertEquals(new DrawingCoordinateReport.Emu(17), report.to().dy());
    assertEquals(before, drawing.xmlText());
  }

  @Test
  void absoluteSignedCoordinatesAndZeroExtentRemainFacts() {
    var anchor = CTDrawing.Factory.newInstance().addNewAbsoluteAnchor();
    anchor.addNewPos().setX(-1L);
    anchor.getPos().setY("1.00in");
    anchor.addNewExt().setCx(0);
    anchor.getExt().setCy(27273042316900L);
    var shape = anchor.addNewSp();
    var report =
        assertInstanceOf(DrawingAnchorReport.Absolute.class, DrawingAdapter.snapshotAnchor(shape));
    assertEquals(new DrawingCoordinateReport.Emu(-1), report.x());
    assertEquals(new DrawingCoordinateReport.Measured("1.00in"), report.y());
    assertEquals(0, report.widthEmu());
    assertEquals(27273042316900L, report.heightEmu());
    anchor.getExt().setCy(27273042316901L);
    assertEquals(
        WorkbookContentIssue.INVALID_NATIVE_VALUE,
        assertThrows(
                WorkbookContentRefusalException.class, () -> DrawingAdapter.snapshotAnchor(shape))
            .issue());
  }

  @Test
  void vmlSeventeenIsPixelsRatherThanDrawingMlEmu() throws IOException {
    try (var input =
            Objects.requireNonNull(
                getClass().getResourceAsStream("/workbooks/signature-line-source.xlsx"));
        var workbook = new XSSFWorkbook(input)) {
      var report = new SignatureAdapter().signatureLines(workbook.getSheetAt(0)).getFirst();
      assertEquals(
          new DrawingCoordinateReport.Pixels(17),
          assertInstanceOf(DrawingAnchorReport.TwoCell.class, report.anchor()).from().dx());
    }
  }

  @Test
  void missingOrInvalidNativeScalarsAreLocatedRefusals() {
    var marker =
        org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker.Factory
            .newInstance();
    var missing =
        assertThrows(WorkbookContentRefusalException.class, () -> DrawingAdapter.marker(marker));
    assertEquals("drawing.anchor.col", missing.location().orElseThrow().fieldName().orElseThrow());
    marker.setCol(0);
    marker.setRow(0);
    marker.setColOff("1zz");
    marker.setRowOff(0L);
    assertEquals(
        WorkbookContentIssue.INVALID_NATIVE_VALUE,
        assertThrows(WorkbookContentRefusalException.class, () -> DrawingAdapter.marker(marker))
            .issue());
  }

  @Test
  void unknownNativeRuntimeIsInternalAndFatalErrorsPropagateByIdentity() {
    var failure = new IllegalArgumentException("synthetic provider invariant");
    var broken = failingMarker(failure);
    assertSame(
        failure,
        assertThrows(UnexpectedNativeContentException.class, () -> DrawingAdapter.marker(broken))
            .getCause());
    var fatal = new AssertionError("synthetic fatal provider error");
    var fatalMarker = failingMarker(fatal);
    assertSame(fatal, assertThrows(AssertionError.class, () -> DrawingAdapter.marker(fatalMarker)));
  }

  // The fault proxy belongs to the native interface's loader, independent of an embedding host
  // TCCL.
  @SuppressWarnings("PMD.UseProperClassLoader")
  private static org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker
      failingMarker(Throwable failure) {
    var type = org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker.class;
    return (org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker)
        Proxy.newProxyInstance(
            type.getClassLoader(),
            new Class<?>[] {type},
            (proxy, method, arguments) -> {
              throw failure;
            });
  }
}
