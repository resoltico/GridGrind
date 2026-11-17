package dev.erst.gridgrind.core.model.value;

import java.time.LocalDate;

/** Workbook-owned serial calendar, independent of host locale and timezone. */
public enum ExcelDateSystem {
  WINDOWS_1900(LocalDate.of(1900, 1, 1), LocalDate.of(1899, 12, 31)),
  MAC_1904(LocalDate.of(1904, 1, 1), LocalDate.of(1904, 1, 1));

  public static final LocalDate MAXIMUM_DATE = LocalDate.of(9999, 12, 31);
  private final LocalDate minimum;
  private final LocalDate origin;

  ExcelDateSystem(LocalDate minimum, LocalDate origin) {
    this.minimum = minimum;
    this.origin = origin;
  }

  /** Returns the earliest civil date supported by this workbook calendar. */
  public LocalDate minimumDate() {
    return minimum;
  }

  LocalDate origin() {
    return origin;
  }

  /** Returns the serial for the last supported civil date, 9999-12-31. */
  public double maximumSerial() {
    return TemporalSemantics.serial(MAXIMUM_DATE, this);
  }
}
