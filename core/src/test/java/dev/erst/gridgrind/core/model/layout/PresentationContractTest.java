package dev.erst.gridgrind.core.model.layout;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsSpec;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.print.PrintSetupSpec;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Validates authored and factual presentation and print contracts. */
class PresentationContractTest {
  @Test
  void sheetPresentationValueObjectsNormalizeConvertAndValidate() {
    IgnoredErrorSpec ignoredError =
        new IgnoredErrorSpec(
            new dev.erst.gridgrind.core.vocabulary.RangeRef("B2:A1"),
            List.of(ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT));
    SheetPresentationSpec presentation =
        new SheetPresentationSpec(
            new SheetDisplaySpec(false, false, false, true, true),
            Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
            new SheetOutlineSummarySpec(false, false),
            new SheetDefaultsSpec(11, 18.5d),
            List.of(ignoredError));
    assertEquals("B2:A1", ignoredError.range().value());
    assertEquals(List.of(ignoredError), presentation.ignoredErrors());

    List<ExcelIgnoredErrorType> errorTypesWithNull =
        new ArrayList<>(Arrays.asList(ExcelIgnoredErrorType.FORMULA, null));
    List<IgnoredErrorSpec> ignoredErrorsWithNull =
        new ArrayList<>(Arrays.asList(ignoredError, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new IgnoredErrorSpec(new dev.erst.gridgrind.core.vocabulary.RangeRef(null), List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new IgnoredErrorSpec(new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"), null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new IgnoredErrorSpec(new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new IgnoredErrorSpec(
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"), errorTypesWithNull));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new IgnoredErrorSpec(
                new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"),
                List.of(ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.FORMULA)));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsSpec(0, 15.0d));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsSpec(8, 0.0d));
    assertThrows(IllegalArgumentException.class, () -> new SheetDefaultsSpec(8, Double.NaN));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationSpec(
                null,
                null,
                SheetOutlineSummarySpec.defaults(),
                SheetDefaultsSpec.defaults(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationSpec(
                SheetDisplaySpec.defaults(), null, null, SheetDefaultsSpec.defaults(), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationSpec(
                SheetDisplaySpec.defaults(),
                null,
                SheetOutlineSummarySpec.defaults(),
                null,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationSpec(
                SheetDisplaySpec.defaults(),
                null,
                SheetOutlineSummarySpec.defaults(),
                SheetDefaultsSpec.defaults(),
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationSpec(
                SheetDisplaySpec.defaults(),
                null,
                SheetOutlineSummarySpec.defaults(),
                SheetDefaultsSpec.defaults(),
                ignoredErrorsWithNull));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationReport(
                null,
                null,
                SheetOutlineSummaryReport.defaults(),
                SheetDefaultsReport.defaults(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                null,
                null,
                SheetDefaultsReport.defaults(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                null,
                SheetOutlineSummaryReport.defaults(),
                null,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                null,
                SheetOutlineSummaryReport.defaults(),
                SheetDefaultsReport.defaults(),
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new SheetPresentationReport(
                SheetDisplayReport.defaults(),
                null,
                SheetOutlineSummaryReport.defaults(),
                SheetDefaultsReport.defaults(),
                new ArrayList<>(
                    Arrays.asList(
                        (dev.erst.gridgrind.core.model.validation.IgnoredErrorReport) null))));
  }

  @Test
  void printSetupValueObjectsValidateIndexCollections() {
    PrintMarginsSpec margins = new PrintMarginsSpec(0.5d, 0.6d, 0.7d, 0.8d, 0.2d, 0.3d);
    PrintMarginsReport marginsSnapshot = new PrintMarginsReport(0.5d, 0.6d, 0.7d, 0.8d, 0.2d, 0.3d);
    PrintSetupSpec setup =
        new PrintSetupSpec(
            margins, true, true, false, 8, true, false, 2, true, 4, List.of(1), List.of(2));
    PrintSetupReport snapshot =
        new PrintSetupReport(
            marginsSnapshot, true, true, false, 8, true, false, 2, true, 4, List.of(1), List.of(2));

    assertTrue(setup.printGridlines());
    assertTrue(snapshot.printGridlines());

    List<Integer> indexesWithNull = new ArrayList<>(Arrays.asList(1, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupSpec(
                null, false, false, false, 1, false, false, 1, false, 1, List.of(), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupSpec(
                margins, false, false, false, 1, false, false, 1, false, 1, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupSpec(
                margins, false, false, false, 1, false, false, 1, false, 1, List.of(), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupSpec(
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
            new PrintSetupReport(
                null, false, false, false, 1, false, false, 1, false, 1, List.of(), List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new PrintSetupReport(
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
            new PrintSetupReport(
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
            new PrintSetupReport(
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
