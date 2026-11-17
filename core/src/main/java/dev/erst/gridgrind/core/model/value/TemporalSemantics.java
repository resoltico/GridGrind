package dev.erst.gridgrind.core.model.value;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;
import java.util.Objects;

/** Pure civil serial conversion and interpretation at the declared millisecond resolution. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_TEMPORAL_CIVIL_DOMAIN,
  dev.erst.gridgrind.core.limits.Limit.ID_TEMPORAL_MILLISECOND_PRECISION
})
public final class TemporalSemantics {
  private static final long MILLIS_PER_DAY = 86_400_000;
  private static final LocalDate LEAP_OFFSET_START = LocalDate.of(1900, 3, 1);

  private TemporalSemantics() {}

  /** Refuses dates outside the shared 1900-01-01 through 9999-12-31 civil domain. */
  public static void requireDateShape(LocalDate value) {
    Objects.requireNonNull(value, "value must not be null");
    if (!supportedDateShape(value))
      throw invalid(
          "DATE",
          dev.erst.gridgrind.core.problem.FieldValidationTemporalRule.DATE_RANGE,
          ExcelDateSystem.WINDOWS_1900.minimumDate().toString(),
          ExcelDateSystem.MAXIMUM_DATE.toString());
  }

  /** Refuses out-of-domain civil dates and fractions finer than a millisecond. */
  public static void requireDateTimeShape(LocalDateTime value) {
    Objects.requireNonNull(value, "value must not be null");
    requireDateShape(value.toLocalDate());
    if (!value.equals(value.truncatedTo(ChronoUnit.MILLIS)))
      throw invalid(
          "DATE_TIME",
          dev.erst.gridgrind.core.problem.FieldValidationTemporalRule.MILLISECOND_PRECISION);
  }

  /** Converts a civil date to a serial, refusing dates before the selected calendar minimum. */
  public static double serial(LocalDate value, ExcelDateSystem system) {
    requireDateShape(value);
    Objects.requireNonNull(system, "system must not be null");
    if (value.isBefore(system.minimumDate()))
      throw invalid(
          "DATE",
          dev.erst.gridgrind.core.problem.FieldValidationTemporalRule.ACTIVE_EPOCH,
          system.minimumDate().toString(),
          system.name());
    long days = value.toEpochDay() - system.origin().toEpochDay();
    if (system == ExcelDateSystem.WINDOWS_1900 && !value.isBefore(LEAP_OFFSET_START)) days++;
    return days;
  }

  /** Converts a millisecond-precision civil date and time using the selected workbook calendar. */
  public static double serial(LocalDateTime value, ExcelDateSystem system) {
    requireDateTimeShape(value);
    return serial(value.toLocalDate(), system)
        + value.toLocalTime().getLong(ChronoField.MILLI_OF_DAY) / (double) MILLIS_PER_DAY;
  }

  /** Validates temporal write operands against the workbook calendar; other writes pass through. */
  public static void requireWritable(ResolvedCellWrite value, ExcelDateSystem system) {
    switch (value) {
      case CellValueSpec.Date date -> serial(date.value(), system);
      case CellValueSpec.DateTime dateTime -> serial(dateTime.value(), system);
      default -> {}
    }
  }

  /**
   * Interprets a numeric fact at millisecond resolution or reports why its temporal form is
   * unrepresentable.
   */
  public static CellTemporalReport interpret(
      double number, CellTemporalKind kind, ExcelDateSystem system) {
    Objects.requireNonNull(kind);
    Objects.requireNonNull(system);
    if (kind == CellTemporalKind.ELAPSED)
      return unrepresentable(kind, CellTemporalReason.UNSUPPORTED_ELAPSED_FORMAT);
    if (!Double.isFinite(number)) return unrepresentable(kind, CellTemporalReason.NONFINITE);
    if (number < 0) return unrepresentable(kind, CellTemporalReason.NEGATIVE);
    if (kind == CellTemporalKind.TIME) {
      long millis = Math.round((number - Math.floor(number)) * MILLIS_PER_DAY) % MILLIS_PER_DAY;
      return new CellTemporalReport.Interpreted(
          kind, LocalTime.ofNanoOfDay(millis * 1_000_000).toString());
    }
    double minimum = system == ExcelDateSystem.WINDOWS_1900 ? 1 : 0;
    if (number < minimum || number >= system.maximumSerial() + 1)
      return unrepresentable(kind, CellTemporalReason.OUTSIDE_CIVIL_RANGE);
    if (system == ExcelDateSystem.WINDOWS_1900 && (long) Math.floor(number) == 60)
      return unrepresentable(kind, CellTemporalReason.FICTITIOUS_1900_LEAP_DAY);
    long millis = Math.round(number * MILLIS_PER_DAY);
    long day = millis / MILLIS_PER_DAY;
    if (day > system.maximumSerial())
      return unrepresentable(kind, CellTemporalReason.OUTSIDE_CIVIL_RANGE);
    if (system == ExcelDateSystem.WINDOWS_1900 && day == 60)
      return unrepresentable(kind, CellTemporalReason.FICTITIOUS_1900_LEAP_DAY);
    long civilDay = system == ExcelDateSystem.WINDOWS_1900 && day > 60 ? day - 1 : day;
    LocalDate date = system.origin().plusDays(civilDay);
    return new CellTemporalReport.Interpreted(
        kind,
        kind == CellTemporalKind.DATE
            ? date.toString()
            : LocalDateTime.of(date, LocalTime.ofNanoOfDay((millis % MILLIS_PER_DAY) * 1_000_000))
                .toString());
  }

  static boolean supportedDateShape(LocalDate value) {
    return !value.isBefore(ExcelDateSystem.WINDOWS_1900.minimumDate())
        && !value.isAfter(ExcelDateSystem.MAXIMUM_DATE);
  }

  static boolean millisecondTime(LocalTime value) {
    return value.equals(value.truncatedTo(ChronoUnit.MILLIS));
  }

  private static dev.erst.gridgrind.core.problem.ValueInvariantException invalid(
      String field,
      dev.erst.gridgrind.core.problem.FieldValidationTemporalRule rule,
      String... operands) {
    return new dev.erst.gridgrind.core.problem.ValueInvariantException(
        dev.erst.gridgrind.core.problem.FieldValidationProblem.atField(field, rule, operands));
  }

  private static CellTemporalReport unrepresentable(
      CellTemporalKind kind, CellTemporalReason reason) {
    return new CellTemporalReport.Unrepresentable(kind, reason);
  }
}
