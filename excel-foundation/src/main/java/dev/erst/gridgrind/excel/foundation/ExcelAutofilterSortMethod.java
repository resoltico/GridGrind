package dev.erst.gridgrind.excel.foundation;

import java.util.Arrays;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Supported OOXML autofilter sort-method values exposed as a typed protocol surface. */
public enum ExcelAutofilterSortMethod {
  PINYIN("pinYin"),
  STROKE("stroke");

  private final String ooxmlValue;

  ExcelAutofilterSortMethod(String ooxmlValue) {
    this.ooxmlValue = ooxmlValue;
  }

  /** Returns the exact OOXML string used by SpreadsheetML for this sort method. */
  public String ooxmlValue() {
    return ooxmlValue;
  }

  /** Resolves one enum constant from the exact OOXML value emitted by SpreadsheetML. */
  public static Optional<ExcelAutofilterSortMethod> fromOoxmlValue(String value) {
    Objects.requireNonNull(value, "value must not be null");
    return Arrays.stream(values()).filter(method -> method.ooxmlValue.equals(value)).findFirst();
  }

  /** Resolves one enum constant from the protocol-facing uppercase token. */
  public static Optional<ExcelAutofilterSortMethod> fromProtocolToken(String value) {
    Objects.requireNonNull(value, "value must not be null");
    String normalized = value.trim().toUpperCase(Locale.ROOT);
    return Arrays.stream(values()).filter(method -> method.name().equals(normalized)).findFirst();
  }
}
