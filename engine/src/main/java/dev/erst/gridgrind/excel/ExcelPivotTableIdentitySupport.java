package dev.erst.gridgrind.excel;

import java.util.Locale;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

/** Identity, naming, and location helpers shared across pivot-table workflows. */
final class ExcelPivotTableIdentitySupport {
  private static final String SYNTHETIC_PREFIX = "_GG_PIVOT_";

  private ExcelPivotTableIdentitySupport() {}

  static PivotLocation safeLocation(PivotHandle handle) {
    String rawRange = rawLocationRange(handle);
    if (rawRange == null) {
      return null;
    }
    try {
      AreaReference area = new AreaReference(rawRange, SpreadsheetVersion.EXCEL2007);
      String locationRange = normalizeArea(area);
      return new PivotLocation(area.getFirstCell().formatAsString(), locationRange);
    } catch (RuntimeException exception) {
      return null;
    }
  }

  static String rawLocationRange(PivotHandle handle) {
    var location = handle.table().getCTPivotTableDefinition().getLocation();
    return location == null ? null : nullIfBlank(location.getRef());
  }

  static String resolvedName(PivotHandle handle) {
    String actualName = actualName(handle);
    if (actualName != null) {
      return actualName;
    }
    PivotLocation location = safeLocation(handle);
    return syntheticName(
        handle.sheetName(),
        location == null ? "PIVOT_" + (handle.ordinalOnSheet() + 1) : location.topLeftAddress());
  }

  static String actualName(PivotHandle handle) {
    return nullIfBlank(handle.table().getCTPivotTableDefinition().getName());
  }

  static String syntheticName(String sheetName, String topLeftAddress) {
    return SYNTHETIC_PREFIX + sanitize(sheetName) + '_' + sanitize(topLeftAddress);
  }

  static String sanitize(String value) {
    StringBuilder builder = new StringBuilder(value.length());
    for (int index = 0; index < value.length(); index++) {
      char character = value.charAt(index);
      if (Character.isLetterOrDigit(character) || character == '_') {
        builder.append(character);
      } else {
        builder.append('_');
      }
    }
    return builder.toString();
  }

  static AreaReference contiguousArea(String rawRange, String fieldName) {
    try {
      return new AreaReference(rawRange, SpreadsheetVersion.EXCEL2007);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          fieldName + " must be a contiguous A1-style range", exception);
    }
  }

  static String normalizeArea(AreaReference area) {
    CellReference first =
        new CellReference(area.getFirstCell().getRow(), area.getFirstCell().getCol(), false, false);
    CellReference last =
        new CellReference(area.getLastCell().getRow(), area.getLastCell().getCol(), false, false);
    boolean singleCell = first.getRow() == last.getRow() && first.getCol() == last.getCol();
    return singleCell
        ? first.formatAsString()
        : first.formatAsString() + ":" + last.formatAsString();
  }

  static String requireNonBlank(String value, String message) {
    if (value == null || value.isBlank()) {
      throw new IllegalArgumentException(message);
    }
    return value;
  }

  static String nonBlankOrDefault(String value, String defaultValue) {
    return value == null || value.isBlank() ? defaultValue : value;
  }

  static String normalizedResolvedName(PivotHandle handle) {
    return resolvedName(handle).toUpperCase(Locale.ROOT);
  }

  private static String nullIfBlank(String value) {
    return value == null || value.isBlank() ? null : value;
  }
}
