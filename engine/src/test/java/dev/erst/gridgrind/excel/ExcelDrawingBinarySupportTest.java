package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject;

/** Focused validation for drawing-binary helper guard rails. */
class ExcelDrawingBinarySupportTest {
  @Test
  void previewSheetRelationIdRejectsBlankValues() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelDrawingBinarySupport.setPreviewSheetRelationId(
                    CTOleObject.Factory.newInstance(), " "));
    assertTrue(failure.getMessage().contains("relationId must not be blank"));
  }
}
