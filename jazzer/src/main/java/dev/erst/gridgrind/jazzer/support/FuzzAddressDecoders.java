package dev.erst.gridgrind.jazzer.support;

import java.util.List;
import java.util.Objects;

/** Owns bounded sheet-name, address, range, and free-text decoding for Jazzer inputs. */
final class FuzzAddressDecoders {
  static final List<String> VALID_COLUMNS = List.of("A", "B", "C", "D", "E", "F", "G", "H");

  private FuzzAddressDecoders() {}

  static String nextSheetName(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (!valid) {
      return switch (data.consumeInt(0, 5)) {
        case 0 -> "";
        case 1 -> " ";
        case 2 -> "Bad/Name";
        case 3 -> "Bad*Name";
        case 4 -> "Bad[Name]";
        default -> "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEF";
      };
    }

    StringBuilder builder = new StringBuilder();
    int length = data.consumeInt(1, 12);
    for (int index = 0; index < length; index++) {
      char value = (char) data.consumeInt('A', 'Z');
      if ("[]:*?/\\'".indexOf(value) >= 0) {
        value = 'S';
      }
      builder.append(value);
    }
    return builder.toString();
  }

  static String nextCellAddress(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (!valid) {
      return switch (data.consumeInt(0, 4)) {
        case 0 -> "";
        case 1 -> "ZZZ999999";
        case 2 -> "A0";
        case 3 -> "R1C1";
        default -> "1A";
      };
    }

    String column = VALID_COLUMNS.get(data.consumeInt(0, VALID_COLUMNS.size() - 1));
    int row = data.consumeInt(1, 25);
    return column + row;
  }

  static String nextNonBlankCellAddress(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (valid) {
      return nextCellAddress(data, true);
    }

    return switch (data.consumeInt(0, 3)) {
      case 0 -> "ZZZ999999";
      case 1 -> "A0";
      case 2 -> "R1C1";
      default -> "1A";
    };
  }

  static String nextRange(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (!valid) {
      return switch (data.consumeInt(0, 4)) {
        case 0 -> "";
        case 1 -> "A1:";
        case 2 -> "A1:1B";
        case 3 -> "A0:B2";
        default -> "B2:A1";
      };
    }

    int firstColumnIndex = data.consumeInt(0, VALID_COLUMNS.size() - 2);
    int lastColumnIndex = data.consumeInt(firstColumnIndex, firstColumnIndex + 1);
    int firstRow = data.consumeInt(1, 6);
    int lastRow = data.consumeInt(firstRow, firstRow + 2);
    return VALID_COLUMNS.get(firstColumnIndex)
        + firstRow
        + ":"
        + VALID_COLUMNS.get(lastColumnIndex)
        + lastRow;
  }

  static String nextNonBlankRange(GridGrindFuzzData data, boolean valid) {
    Objects.requireNonNull(data, "data must not be null");

    if (valid) {
      return nextRange(data, true);
    }

    return switch (data.consumeInt(0, 3)) {
      case 0 -> "A1:";
      case 1 -> "A1:1B";
      case 2 -> "A0:B2";
      default -> "B2:A1";
    };
  }

  static String nextText(GridGrindFuzzData data) {
    int length = data.consumeInt(1, 16);
    StringBuilder builder = new StringBuilder(length);
    for (int index = 0; index < length; index++) {
      builder.append((char) data.consumeInt('A', 'z'));
    }
    String result = builder.toString().trim();
    return result.isBlank() ? "X" : result;
  }

  static String nextFormula(GridGrindFuzzData data) {
    return switch (data.consumeInt(0, 5)) {
      case 0 -> "SUM(A1:A2)";
      case 1 -> "A1+A2";
      case 2 -> "1/0";
      case 3 -> "IF(A1>0,1,0)";
      case 4 -> "BAD(";
      default -> nextText(data);
    };
  }
}
