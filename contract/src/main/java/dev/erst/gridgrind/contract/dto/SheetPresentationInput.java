package dev.erst.gridgrind.contract.dto;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/** Authoritative sheet-presentation payload for display flags, defaults, and ignored errors. */
public record SheetPresentationInput(
    SheetDisplayInput display,
    ColorInput tabColor,
    SheetOutlineSummaryInput outlineSummary,
    SheetDefaultsInput sheetDefaults,
    List<IgnoredErrorInput> ignoredErrors) {
  /** Returns the effective default sheet-presentation payload for one new worksheet. */
  public static SheetPresentationInput defaults() {
    return new SheetPresentationInput(
        SheetDisplayInput.defaults(),
        null,
        SheetOutlineSummaryInput.defaults(),
        SheetDefaultsInput.defaults(),
        List.of());
  }

  public SheetPresentationInput {
    display = display == null ? defaults().display() : display;
    outlineSummary = outlineSummary == null ? defaults().outlineSummary() : outlineSummary;
    sheetDefaults = sheetDefaults == null ? defaults().sheetDefaults() : sheetDefaults;
    ignoredErrors =
        List.copyOf(
            Objects.requireNonNull(
                ignoredErrors == null ? List.of() : ignoredErrors,
                "ignoredErrors must not be null"));
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(outlineSummary, "outlineSummary must not be null");
    Objects.requireNonNull(sheetDefaults, "sheetDefaults must not be null");
    List<String> normalizedRanges = ignoredErrors.stream().map(IgnoredErrorInput::range).toList();
    if (normalizedRanges.size() != new LinkedHashSet<>(normalizedRanges).size()) {
      throw new IllegalArgumentException("ignoredErrors must not contain duplicate ranges");
    }
  }
}
