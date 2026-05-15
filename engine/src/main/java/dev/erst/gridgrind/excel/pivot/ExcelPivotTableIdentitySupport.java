package dev.erst.gridgrind.excel.pivot;

import java.util.Locale;
import java.util.Optional;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

/** Identity, naming, and location helpers shared across pivot-table workflows. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelPivotTableIdentitySupport {
  private static final String SYNTHETIC_PREFIX = "_GG_PIVOT_";

  private ExcelPivotTableIdentitySupport() {}

  public static Optional<PivotLocation> safeLocation(PivotHandle handle) {
    Optional<String> rawRange = rawLocationRange(handle);
    if (rawRange.isEmpty()) {
      return Optional.empty();
    }
    try {
      AreaReference area = new AreaReference(rawRange.orElseThrow(), SpreadsheetVersion.EXCEL2007);
      String locationRange = normalizeArea(area);
      return Optional.of(new PivotLocation(area.getFirstCell().formatAsString(), locationRange));
    } catch (RuntimeException exception) {
      return Optional.empty();
    }
  }

  public static Optional<String> rawLocationRange(PivotHandle handle) {
    var location = handle.table().getCTPivotTableDefinition().getLocation();
    return location == null ? Optional.empty() : blankAsOptional(location.getRef());
  }

  public static String resolvedName(PivotHandle handle) {
    return actualName(handle)
        .orElseGet(
            () ->
                syntheticName(
                    handle.sheetName(),
                    safeLocation(handle)
                        .map(PivotLocation::topLeftAddress)
                        .orElse("PIVOT_" + (handle.ordinalOnSheet() + 1))));
  }

  public static Optional<String> actualName(PivotHandle handle) {
    return blankAsOptional(handle.table().getCTPivotTableDefinition().getName());
  }

  public static String syntheticName(String sheetName, String topLeftAddress) {
    return SYNTHETIC_PREFIX + sanitize(sheetName) + '_' + sanitize(topLeftAddress);
  }

  public static String sanitize(String value) {
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

  public static AreaReference contiguousArea(String rawRange, String fieldName) {
    try {
      return new AreaReference(rawRange, SpreadsheetVersion.EXCEL2007);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          fieldName + " must be a contiguous A1-style range", exception);
    }
  }

  public static String normalizeArea(AreaReference area) {
    CellReference first =
        new CellReference(area.getFirstCell().getRow(), area.getFirstCell().getCol(), false, false);
    CellReference last =
        new CellReference(area.getLastCell().getRow(), area.getLastCell().getCol(), false, false);
    boolean singleCell = first.getRow() == last.getRow() && first.getCol() == last.getCol();
    return singleCell
        ? first.formatAsString()
        : first.formatAsString() + ":" + last.formatAsString();
  }

  public static String requireNonBlank(String value, String message) {
    if (value == null || value.isBlank()) {
      throw new IllegalArgumentException(message);
    }
    return value;
  }

  public static String nonBlankOrDefault(String value, String defaultValue) {
    return value == null || value.isBlank() ? defaultValue : value;
  }

  static String normalizedResolvedName(PivotHandle handle) {
    return resolvedName(handle).toUpperCase(Locale.ROOT);
  }

  private static Optional<String> blankAsOptional(String value) {
    return value == null || value.isBlank() ? Optional.empty() : Optional.of(value);
  }
}
