package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
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

  @Test
  void sheetPresentationValueObjectsNormalizeConvertAndValidate() {
    ExcelIgnoredError ignoredError =
        new ExcelIgnoredError(
            "B2:A1",
            List.of(ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT));
    ExcelSheetPresentation presentation =
        new ExcelSheetPresentation(
            new ExcelSheetDisplay(false, false, false, true, true),
            new ExcelColor("#112233"),
            new ExcelSheetOutlineSummary(false, false),
            new ExcelSheetDefaults(11, 18.5d),
            List.of(ignoredError));
    ExcelSheetPresentationSnapshot snapshot =
        new ExcelSheetPresentationSnapshot(
            presentation.display(),
            new ExcelColorSnapshot("#112233"),
            presentation.outlineSummary(),
            presentation.sheetDefaults(),
            presentation.ignoredErrors());

    assertEquals("A1:B2", ignoredError.range());
    assertEquals(presentation, snapshot.toAuthoringPresentation());
    assertNull(
        new ExcelSheetPresentationSnapshot(
                ExcelSheetDisplay.defaults(),
                null,
                ExcelSheetOutlineSummary.defaults(),
                ExcelSheetDefaults.defaults(),
                List.of())
            .toAuthoringPresentation()
            .tabColor());

    List<ExcelIgnoredErrorType> errorTypesWithNull =
        new ArrayList<>(Arrays.asList(ExcelIgnoredErrorType.FORMULA, null));
    List<ExcelIgnoredError> ignoredErrorsWithNull =
        new ArrayList<>(Arrays.asList(ignoredError, null));
    assertThrows(NullPointerException.class, () -> new ExcelIgnoredError(null, List.of()));
    assertThrows(NullPointerException.class, () -> new ExcelIgnoredError("A1", null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelIgnoredError("A1", List.of()));
    assertThrows(NullPointerException.class, () -> new ExcelIgnoredError("A1", errorTypesWithNull));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelIgnoredError(
                "A1", List.of(ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.FORMULA)));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetDefaults(0, 15.0d));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetDefaults(8, 0.0d));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetDefaults(8, Double.NaN));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentation(
                null,
                null,
                ExcelSheetOutlineSummary.defaults(),
                ExcelSheetDefaults.defaults(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentation(
                ExcelSheetDisplay.defaults(),
                null,
                null,
                ExcelSheetDefaults.defaults(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentation(
                ExcelSheetDisplay.defaults(),
                null,
                ExcelSheetOutlineSummary.defaults(),
                null,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentation(
                ExcelSheetDisplay.defaults(),
                null,
                ExcelSheetOutlineSummary.defaults(),
                ExcelSheetDefaults.defaults(),
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentation(
                ExcelSheetDisplay.defaults(),
                null,
                ExcelSheetOutlineSummary.defaults(),
                ExcelSheetDefaults.defaults(),
                ignoredErrorsWithNull));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentationSnapshot(
                null,
                null,
                ExcelSheetOutlineSummary.defaults(),
                ExcelSheetDefaults.defaults(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentationSnapshot(
                ExcelSheetDisplay.defaults(),
                null,
                null,
                ExcelSheetDefaults.defaults(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentationSnapshot(
                ExcelSheetDisplay.defaults(),
                null,
                ExcelSheetOutlineSummary.defaults(),
                null,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentationSnapshot(
                ExcelSheetDisplay.defaults(),
                null,
                ExcelSheetOutlineSummary.defaults(),
                ExcelSheetDefaults.defaults(),
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelSheetPresentationSnapshot(
                ExcelSheetDisplay.defaults(),
                null,
                ExcelSheetOutlineSummary.defaults(),
                ExcelSheetDefaults.defaults(),
                ignoredErrorsWithNull));
  }

  @Test
  void printSetupValueObjectsValidateIndexCollections() {
    ExcelPrintMargins margins = new ExcelPrintMargins(0.5d, 0.6d, 0.7d, 0.8d, 0.2d, 0.3d);
    ExcelPrintMarginsSnapshot marginsSnapshot =
        new ExcelPrintMarginsSnapshot(0.5d, 0.6d, 0.7d, 0.8d, 0.2d, 0.3d);
    ExcelPrintSetup setup =
        new ExcelPrintSetup(
            margins, true, true, false, 8, true, false, 2, true, 4, List.of(1), List.of(2));
    ExcelPrintSetupSnapshot snapshot =
        new ExcelPrintSetupSnapshot(
            marginsSnapshot, true, true, false, 8, true, false, 2, true, 4, List.of(1), List.of(2));

    assertTrue(setup.printGridlines());
    assertTrue(snapshot.printGridlines());

    List<Integer> indexesWithNull = new ArrayList<>(Arrays.asList(1, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetup(
                null, false, false, false, 1, false, false, 1, false, 1, List.of(), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetup(
                margins, false, false, false, 1, false, false, 1, false, 1, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetup(
                margins, false, false, false, 1, false, false, 1, false, 1, List.of(), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetup(
                margins,
                false,
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                1,
                indexesWithNull,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                null, false, false, false, 1, false, false, 1, false, 1, List.of(), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                marginsSnapshot,
                false,
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                1,
                null,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                marginsSnapshot,
                false,
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                1,
                List.of(),
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPrintSetupSnapshot(
                marginsSnapshot,
                false,
                false,
                false,
                1,
                false,
                false,
                1,
                false,
                1,
                indexesWithNull,
                List.of()));
  }
}
