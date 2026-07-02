package dev.erst.gridgrind.excel;

import java.util.Locale;
import java.util.Optional;
import org.apache.poi.ss.usermodel.DateUtil;
import org.jspecify.annotations.Nullable;

/** Shared number-format heuristics for projected temporal readback metadata. */
public final class ExcelTemporalFormatSupport {
  /** Temporal semantic inferred from a numeric cell's observed number format. */
  public enum ObservedKind {
    DATE,
    TIME,
    DATE_TIME
  }

  private ExcelTemporalFormatSupport() {}

  /** Returns the temporal semantic implied by one Excel number format, when any. */
  public static Optional<ObservedKind> observedKind(String numberFormat) {
    if (!DateUtil.isADateFormat(-1, numberFormat)) {
      return Optional.empty();
    }
    String normalized = normalizeFormat(numberFormat);
    boolean containsTimeFields = containsTimeFields(normalized);
    if (!containsTimeFields) {
      return Optional.of(ObservedKind.DATE);
    }
    return Optional.of(containsDateFields(normalized) ? ObservedKind.DATE_TIME : ObservedKind.TIME);
  }

  private static boolean containsTimeFields(String normalizedFormat) {
    return normalizedFormat.contains("am/pm")
        || normalizedFormat.contains("a/p")
        || normalizedFormat.contains(":")
        || normalizedFormat.contains("h")
        || normalizedFormat.contains("s");
  }

  private static boolean containsDateFields(String normalizedFormat) {
    return normalizedFormat.contains("y")
        || normalizedFormat.contains("d")
        || containsMonthDateToken(normalizedFormat);
  }

  private static boolean containsMonthDateToken(String normalizedFormat) {
    int cursor = 0;
    while (cursor < normalizedFormat.length()) {
      if (normalizedFormat.charAt(cursor) != 'm') {
        cursor++;
        continue;
      }
      int end = cursor + 1;
      while (end < normalizedFormat.length() && normalizedFormat.charAt(end) == 'm') {
        end++;
      }
      if (!isMinuteToken(normalizedFormat, cursor, end)) {
        return true;
      }
      cursor = end;
    }
    return false;
  }

  private static boolean isMinuteToken(
      String normalizedFormat, int startInclusive, int endExclusive) {
    char previous = previousSignificant(normalizedFormat, startInclusive - 1);
    char next = nextSignificant(normalizedFormat, endExclusive);
    return isTimeContext(previous) || isTimeContext(next);
  }

  private static boolean isTimeContext(char value) {
    return value == 'h' || value == 's' || value == ':';
  }

  private static char previousSignificant(String text, int cursor) {
    int current = cursor;
    while (current >= 0 && Character.isWhitespace(text.charAt(current))) {
      current--;
    }
    return current >= 0 ? text.charAt(current) : '\0';
  }

  private static char nextSignificant(String text, int cursor) {
    int current = cursor;
    while (current < text.length() && Character.isWhitespace(text.charAt(current))) {
      current++;
    }
    return current < text.length() ? text.charAt(current) : '\0';
  }

  private static String normalizeFormat(String numberFormat) {
    StringBuilder builder = new StringBuilder(numberFormat.length());
    boolean quoted = false;
    int cursor = 0;
    while (cursor < numberFormat.length()) {
      char current = lowercased(numberFormat, cursor);
      if (current == '"') {
        quoted = !quoted;
        cursor++;
        continue;
      }
      if (quoted) {
        cursor++;
        continue;
      }
      if (isEscapedCharacter(numberFormat, current, cursor)) {
        cursor += 2;
        continue;
      }
      BracketSection section = readBracketSection(numberFormat, cursor);
      if (section != null) {
        builder.append(section.text());
        cursor = section.nextIndex();
        continue;
      }
      builder.append(current);
      cursor++;
    }
    return builder.toString();
  }

  private static boolean isEscapedCharacter(String numberFormat, char current, int cursor) {
    return current == '\\' && cursor + 1 < numberFormat.length();
  }

  private static char lowercased(String text, int index) {
    return Character.toLowerCase(text.charAt(index));
  }

  private static @Nullable BracketSection readBracketSection(String numberFormat, int cursor) {
    if (lowercased(numberFormat, cursor) != '[') {
      return null;
    }
    int closing = numberFormat.indexOf(']', cursor + 1);
    if (closing < 0) {
      return null;
    }
    String bracket = numberFormat.substring(cursor + 1, closing).toLowerCase(Locale.ROOT);
    return new BracketSection(bracketContainsTime(bracket) ? bracket : "", closing + 1);
  }

  private static boolean bracketContainsTime(String bracket) {
    return bracket.contains("h") || bracket.contains("s");
  }

  private record BracketSection(String text, int nextIndex) {}
}
