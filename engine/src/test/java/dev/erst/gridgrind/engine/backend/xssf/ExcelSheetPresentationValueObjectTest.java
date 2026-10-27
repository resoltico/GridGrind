package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.engine.layout.LayoutAdapter;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.junit.jupiter.api.Test;

/** Focused value-object tests for sheet-presentation and print-setup contracts. */
class ExcelSheetPresentationValueObjectTest {
  @Test
  void ignoredErrorTypesRoundTripEveryPoiFamily() {
    for (ExcelIgnoredErrorType ignoredErrorType : ExcelIgnoredErrorType.values()) {
      assertEquals(
          ignoredErrorType,
          LayoutAdapter.fromPoiIgnoredError(LayoutAdapter.toPoiIgnoredError(ignoredErrorType)));
    }

    for (IgnoredErrorType poiIgnoredErrorType : IgnoredErrorType.values()) {
      assertEquals(
          poiIgnoredErrorType,
          LayoutAdapter.toPoiIgnoredError(LayoutAdapter.fromPoiIgnoredError(poiIgnoredErrorType)));
    }
  }
}
