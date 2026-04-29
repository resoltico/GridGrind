package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
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
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(outlineSummary, "outlineSummary must not be null");
    Objects.requireNonNull(sheetDefaults, "sheetDefaults must not be null");
    ignoredErrors =
        List.copyOf(Objects.requireNonNull(ignoredErrors, "ignoredErrors must not be null"));
    List<String> normalizedRanges = ignoredErrors.stream().map(IgnoredErrorInput::range).toList();
    if (normalizedRanges.size() != new LinkedHashSet<>(normalizedRanges).size()) {
      throw new IllegalArgumentException("ignoredErrors must not contain duplicate ranges");
    }
  }

  @JsonCreator
  static SheetPresentationInput create(
      @JsonProperty("display") SheetDisplayInput display,
      @JsonProperty("tabColor") ColorInput tabColor,
      @JsonProperty("outlineSummary") SheetOutlineSummaryInput outlineSummary,
      @JsonProperty("sheetDefaults") SheetDefaultsInput sheetDefaults,
      @JsonProperty("ignoredErrors") List<IgnoredErrorInput> ignoredErrors) {
    SheetPresentationInput defaults = defaults();
    return new SheetPresentationInput(
        display == null ? defaults.display() : display,
        tabColor,
        outlineSummary == null ? defaults.outlineSummary() : outlineSummary,
        sheetDefaults == null ? defaults.sheetDefaults() : sheetDefaults,
        ignoredErrors == null ? List.of() : ignoredErrors);
  }
}
