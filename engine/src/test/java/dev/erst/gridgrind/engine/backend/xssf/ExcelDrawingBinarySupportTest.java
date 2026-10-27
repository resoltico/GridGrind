package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject;

/** Focused validation for drawing-binary helper guard rails. */
class ExcelDrawingBinarySupportTest {
  @Test
  void previewSheetRelationIdRejectsBlankValues() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> DrawingAdapter.setPreviewSheetRelationId(CTOleObject.Factory.newInstance(), " "));
    assertTrue(failure.getMessage().contains("relationId must not be blank"));
  }
}
