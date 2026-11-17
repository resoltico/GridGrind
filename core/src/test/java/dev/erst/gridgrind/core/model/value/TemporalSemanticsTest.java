package dev.erst.gridgrind.core.model.value;

import static org.junit.jupiter.api.Assertions.*;

import java.time.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Independently specified civil/calendar constants, not values derived from the converter. */
class TemporalSemanticsTest {
  @ParameterizedTest
  @CsvSource({
    "WINDOWS_1900,1900-01-01,1",
    "WINDOWS_1900,1900-02-28,59",
    "WINDOWS_1900,1900-03-01,61",
    "WINDOWS_1900,1904-01-01,1462",
    "WINDOWS_1900,9999-12-31,2958465",
    "MAC_1904,1904-01-01,0",
    "MAC_1904,9999-12-31,2957003"
  })
  void serialAndInterpretationRespectActualEpoch(String epoch, String date, double serial) {
    var system = ExcelDateSystem.valueOf(epoch);
    assertEquals(serial, TemporalSemantics.serial(LocalDate.parse(date), system));
    assertEquals(
        new CellTemporalReport.Interpreted(CellTemporalKind.DATE, date),
        TemporalSemantics.interpret(serial, CellTemporalKind.DATE, system));
  }

  @Test
  void authoringRefusesBoundsAndFinerPrecisionAcrossCompactCarriers() {
    assertThrows(
        IllegalArgumentException.class, () -> new CellValueSpec.Date(LocalDate.of(1880, 1, 1)));
    assertThrows(
        IllegalArgumentException.class,
        () -> TemporalSemantics.serial(LocalDate.of(1903, 12, 31), ExcelDateSystem.MAC_1904));
    var precise = LocalDateTime.parse("2026-11-16T12:34:56.123456789");
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.DateTime(precise));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellGridSpec.DateTimeRows(java.util.List.of(java.util.List.of(precise))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellRowSpec.DateTimeValues(java.util.List.of(precise)));
    assertThrows(
        IllegalArgumentException.class, () -> new CellValueSpec.Date(LocalDate.of(10000, 1, 1)));
  }

  @ParameterizedTest
  @CsvSource({
    "WINDOWS_1900,2026-11-16T12:34:56.123",
    "WINDOWS_1900,9999-12-31T23:59:59.999",
    "MAC_1904,2026-11-16T12:34:56.123",
    "MAC_1904,9999-12-31T23:59:59.999"
  })
  void supportedMillisecondsRoundTripIncludingLastDay(ExcelDateSystem system, String text) {
    double serial = TemporalSemantics.serial(LocalDateTime.parse(text), system);
    assertEquals(
        new CellTemporalReport.Interpreted(CellTemporalKind.DATE_TIME, text),
        TemporalSemantics.interpret(serial, CellTemporalKind.DATE_TIME, system));
  }

  @Test
  void millisecondSerialMatchesIndependentConstant() {
    assertEquals(
        46342.52426068287,
        TemporalSemantics.serial(
            LocalDateTime.parse("2026-11-16T12:34:56.123"), ExcelDateSystem.WINDOWS_1900));
  }

  @Test
  void rawAndRoundedUnrepresentabilityNeverFabricatesAnIsoDate() {
    assertReason(
        -1, CellTemporalKind.DATE, ExcelDateSystem.WINDOWS_1900, CellTemporalReason.NEGATIVE);
    assertReason(
        Double.NaN,
        CellTemporalKind.DATE,
        ExcelDateSystem.WINDOWS_1900,
        CellTemporalReason.NONFINITE);
    assertReason(
        Double.POSITIVE_INFINITY,
        CellTemporalKind.TIME,
        ExcelDateSystem.MAC_1904,
        CellTemporalReason.NONFINITE);
    assertReason(
        0,
        CellTemporalKind.DATE,
        ExcelDateSystem.WINDOWS_1900,
        CellTemporalReason.OUTSIDE_CIVIL_RANGE);
    for (double value : new double[] {60, 60.5, Math.nextDown(60d), Math.nextDown(61d)})
      assertReason(
          value,
          CellTemporalKind.DATE_TIME,
          ExcelDateSystem.WINDOWS_1900,
          CellTemporalReason.FICTITIOUS_1900_LEAP_DAY);
    for (var system : ExcelDateSystem.values()) {
      assertReason(
          system.maximumSerial() + 1,
          CellTemporalKind.DATE,
          system,
          CellTemporalReason.OUTSIDE_CIVIL_RANGE);
      assertReason(
          Math.nextDown(system.maximumSerial() + 1),
          CellTemporalKind.DATE_TIME,
          system,
          CellTemporalReason.OUTSIDE_CIVIL_RANGE);
    }
  }

  @Test
  void clockFractionsDoNotInventACivilDateOrWrapElapsedDuration() {
    var noon = new CellTemporalReport.Interpreted(CellTemporalKind.TIME, "12:00");
    for (double number : new double[] {.5, 60.5, 2958466.5})
      assertEquals(
          noon,
          TemporalSemantics.interpret(number, CellTemporalKind.TIME, ExcelDateSystem.WINDOWS_1900));
    assertEquals(
        new CellTemporalReport.Interpreted(CellTemporalKind.TIME, "00:00"),
        TemporalSemantics.interpret(0, CellTemporalKind.TIME, ExcelDateSystem.WINDOWS_1900));
    assertReason(
        1.5,
        CellTemporalKind.ELAPSED,
        ExcelDateSystem.WINDOWS_1900,
        CellTemporalReason.UNSUPPORTED_ELAPSED_FORMAT);
  }

  @Test
  void malformedInterpretationRetainsItsParseCause() {
    var failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new CellTemporalReport.Interpreted(CellTemporalKind.DATE, "invalid date"));
    assertInstanceOf(java.time.format.DateTimeParseException.class, failure.getCause());
  }

  private static void assertReason(
      double number, CellTemporalKind kind, ExcelDateSystem system, CellTemporalReason reason) {
    var report =
        assertInstanceOf(
            CellTemporalReport.Unrepresentable.class,
            TemporalSemantics.interpret(number, kind, system));
    assertEquals(kind, report.kind());
    assertEquals(reason, report.reason());
  }
}
