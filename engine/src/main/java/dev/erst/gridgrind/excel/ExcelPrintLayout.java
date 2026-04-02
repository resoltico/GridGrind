package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;

/** Supported print-layout state for one sheet. */
public record ExcelPrintLayout(
    Area printArea,
    ExcelPrintOrientation orientation,
    Scaling scaling,
    TitleRows repeatingRows,
    TitleColumns repeatingColumns,
    ExcelHeaderFooterText header,
    ExcelHeaderFooterText footer) {
  /** Returns the GridGrind default print layout for a sheet with no explicit print settings. */
  public static ExcelPrintLayout defaults() {
    return new ExcelPrintLayout(
        new Area.None(),
        ExcelPrintOrientation.PORTRAIT,
        new Scaling.Automatic(),
        new TitleRows.None(),
        new TitleColumns.None(),
        ExcelHeaderFooterText.blank(),
        ExcelHeaderFooterText.blank());
  }

  public ExcelPrintLayout {
    Objects.requireNonNull(printArea, "printArea must not be null");
    Objects.requireNonNull(orientation, "orientation must not be null");
    Objects.requireNonNull(scaling, "scaling must not be null");
    Objects.requireNonNull(repeatingRows, "repeatingRows must not be null");
    Objects.requireNonNull(repeatingColumns, "repeatingColumns must not be null");
    Objects.requireNonNull(header, "header must not be null");
    Objects.requireNonNull(footer, "footer must not be null");
  }

  /** Print-area state for one sheet. */
  public sealed interface Area permits Area.None, Area.Range {
    /** Sheet has no explicit print area. */
    record None() implements Area {}

    /** Sheet prints the provided rectangular A1-style range. */
    record Range(String range) implements Area {
      public Range {
        Objects.requireNonNull(range, "range must not be null");
        if (range.isBlank()) {
          throw new IllegalArgumentException("range must not be blank");
        }
      }
    }
  }

  /** Print scaling state for one sheet. */
  public sealed interface Scaling permits Scaling.Automatic, Scaling.Fit {
    /** Sheet uses Excel's default scaling instead of fit-to-page counts. */
    record Automatic() implements Scaling {}

    /**
     * Sheet fits printed content into the provided page counts.
     *
     * <p>A value of {@code 0} on one axis keeps that axis unconstrained, matching Excel's fit
     * semantics.
     */
    record Fit(int widthPages, int heightPages) implements Scaling {
      public Fit {
        requirePageCount(widthPages, "widthPages");
        requirePageCount(heightPages, "heightPages");
      }
    }
  }

  /** Repeating print-title rows for one sheet. */
  public sealed interface TitleRows permits TitleRows.None, TitleRows.Band {
    /** Sheet has no repeating print-title rows. */
    record None() implements TitleRows {}

    /** Sheet repeats the provided inclusive zero-based row band on every printed page. */
    record Band(int firstRowIndex, int lastRowIndex) implements TitleRows {
      public Band {
        requireBand(firstRowIndex, lastRowIndex, "firstRowIndex", "lastRowIndex");
        if (lastRowIndex > SpreadsheetVersion.EXCEL2007.getLastRowIndex()) {
          throw new IllegalArgumentException(
              "lastRowIndex must not exceed "
                  + SpreadsheetVersion.EXCEL2007.getLastRowIndex()
                  + ": "
                  + lastRowIndex);
        }
      }
    }
  }

  /** Repeating print-title columns for one sheet. */
  public sealed interface TitleColumns permits TitleColumns.None, TitleColumns.Band {
    /** Sheet has no repeating print-title columns. */
    record None() implements TitleColumns {}

    /** Sheet repeats the provided inclusive zero-based column band on every printed page. */
    record Band(int firstColumnIndex, int lastColumnIndex) implements TitleColumns {
      public Band {
        requireBand(firstColumnIndex, lastColumnIndex, "firstColumnIndex", "lastColumnIndex");
        if (lastColumnIndex > SpreadsheetVersion.EXCEL2007.getLastColumnIndex()) {
          throw new IllegalArgumentException(
              "lastColumnIndex must not exceed "
                  + SpreadsheetVersion.EXCEL2007.getLastColumnIndex()
                  + ": "
                  + lastColumnIndex);
        }
      }
    }
  }

  private static void requirePageCount(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
    if (value > Short.MAX_VALUE) {
      throw new IllegalArgumentException(
          fieldName + " must not exceed " + Short.MAX_VALUE + ": " + value);
    }
  }

  private static void requireBand(
      int firstIndex, int lastIndex, String firstFieldName, String lastFieldName) {
    if (firstIndex < 0) {
      throw new IllegalArgumentException(firstFieldName + " must not be negative");
    }
    if (lastIndex < 0) {
      throw new IllegalArgumentException(lastFieldName + " must not be negative");
    }
    if (lastIndex < firstIndex) {
      throw new IllegalArgumentException(
          lastFieldName + " must not be less than " + firstFieldName);
    }
  }
}
