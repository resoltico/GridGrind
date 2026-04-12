package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for the POI relation-removal seam used by drawing and chart deletion. */
class PoiRelationRemovalTest {
  @Test
  void defaultRemoverUsesPoiPrivateRelationRemoval() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));

      assertTrue(PoiRelationRemoval.defaultRemover().test(drawing, chart));
      assertFalse(drawing.getRelations().contains(chart));
    }
  }

  @Test
  void relationHandleLookupFailsCleanlyWithoutPrivateAccess() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> PoiRelationRemoval.removePoiRelationMethod(MethodHandles.publicLookup()));
    assertTrue(failure.getMessage().contains("Failed to resolve POI removeRelation handle"));
  }

  @Test
  void invocationFailuresAreWrappedInProductOwnedMessages() throws Throwable {
    MethodHandle explodingHandle =
        MethodHandles.lookup()
            .findStatic(
                PoiRelationRemovalTest.class,
                "explode",
                MethodType.methodType(
                    boolean.class,
                    POIXMLDocumentPart.class,
                    POIXMLDocumentPart.class,
                    boolean.class));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 6, 10));

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () -> PoiRelationRemoval.invokePoiRelationRemoval(explodingHandle, drawing, chart));
      assertTrue(failure.getMessage().contains("Failed to remove POI relation"));
      assertEquals("boom", failure.getCause().getMessage());
    }
  }

  static boolean explode(
      POIXMLDocumentPart parent, POIXMLDocumentPart child, boolean removeUnusedParts) {
    assertNotNull(parent);
    assertNotNull(child);
    assertTrue(removeUnusedParts);
    throw new IllegalStateException("boom");
  }
}
