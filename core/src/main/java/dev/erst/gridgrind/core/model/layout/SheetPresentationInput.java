package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.model.style.ColorInput;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorInput;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Authoritative sheet-presentation payload for display flags, defaults, and ignored errors. */
@CatalogType(
    id = "SheetPresentationInput",
    summary =
        "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request. Supply explicit display, tabColor, outlineSummary, sheetDefaults, and ignoredErrors values.",
    displayOrder = 69)
public record SheetPresentationInput(
    SheetDisplayInput display,
    Optional<ColorInput> tabColor,
    SheetOutlineSummaryInput outlineSummary,
    SheetDefaultsInput sheetDefaults,
    List<IgnoredErrorInput> ignoredErrors) {
  /** Returns the effective default sheet-presentation payload for one new worksheet. */
  public static SheetPresentationInput defaults() {
    return new SheetPresentationInput(
        SheetDisplayInput.defaults(),
        Optional.empty(),
        SheetOutlineSummaryInput.defaults(),
        SheetDefaultsInput.defaults(),
        List.of());
  }

  public SheetPresentationInput {
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(tabColor, "tabColor must not be null");
    Objects.requireNonNull(outlineSummary, "outlineSummary must not be null");
    Objects.requireNonNull(sheetDefaults, "sheetDefaults must not be null");
    ignoredErrors =
        List.copyOf(Objects.requireNonNull(ignoredErrors, "ignoredErrors must not be null"));
    List<dev.erst.gridgrind.core.vocabulary.RangeRef> normalizedRanges =
        ignoredErrors.stream().map(IgnoredErrorInput::range).toList();
    if (normalizedRanges.size() != new LinkedHashSet<>(normalizedRanges).size()) {
      throw new IllegalArgumentException("ignoredErrors must not contain duplicate ranges");
    }
  }
}
