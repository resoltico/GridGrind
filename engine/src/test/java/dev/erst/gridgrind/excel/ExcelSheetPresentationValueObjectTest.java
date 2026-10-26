package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.junit.jupiter.api.Test;

/** Focused value-object tests for sheet-presentation and print-setup contracts. */
class ExcelSheetPresentationValueObjectTest {
  @Test
  void ignoredErrorTypesRoundTripEveryPoiFamily() {
    for (ExcelIgnoredErrorType ignoredErrorType : ExcelIgnoredErrorType.values()) {
      assertEquals(
          ignoredErrorType,
          ExcelIgnoredErrorPoiBridge.fromPoi(ExcelIgnoredErrorPoiBridge.toPoi(ignoredErrorType)));
    }

    for (IgnoredErrorType poiIgnoredErrorType : IgnoredErrorType.values()) {
      assertEquals(
          poiIgnoredErrorType,
          ExcelIgnoredErrorPoiBridge.toPoi(
              ExcelIgnoredErrorPoiBridge.fromPoi(poiIgnoredErrorType)));
    }
  }
}
