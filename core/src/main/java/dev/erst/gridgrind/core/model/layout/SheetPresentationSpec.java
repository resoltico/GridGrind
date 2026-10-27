package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Authoritative sheet-presentation payload for display flags, defaults, and ignored errors. */
@CatalogType(
    id = "SheetPresentationSpec",
    summary =
        "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request. Supply explicit display, tabColor, outlineSummary, sheetDefaults, and ignoredErrors values.",
    displayOrder = 69)
public record SheetPresentationSpec(
    SheetDisplaySpec display,
    Optional<ColorSpec> tabColor,
    SheetOutlineSummarySpec outlineSummary,
    SheetDefaultsSpec sheetDefaults,
    List<IgnoredErrorSpec> ignoredErrors) {
  /** Returns the effective default sheet-presentation payload for one new worksheet. */
  public static SheetPresentationSpec defaults() {
    return new SheetPresentationSpec(
        SheetDisplaySpec.defaults(),
        Optional.empty(),
        SheetOutlineSummarySpec.defaults(),
        SheetDefaultsSpec.defaults(),
        List.of());
  }

  public SheetPresentationSpec {
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(tabColor, "tabColor must not be null");
    Objects.requireNonNull(outlineSummary, "outlineSummary must not be null");
    Objects.requireNonNull(sheetDefaults, "sheetDefaults must not be null");
    ignoredErrors =
        List.copyOf(Objects.requireNonNull(ignoredErrors, "ignoredErrors must not be null"));
    List<dev.erst.gridgrind.core.vocabulary.RangeRef> normalizedRanges =
        ignoredErrors.stream().map(IgnoredErrorSpec::range).toList();
    if (normalizedRanges.size() != new LinkedHashSet<>(normalizedRanges).size()) {
      throw new IllegalArgumentException("ignoredErrors must not contain duplicate ranges");
    }
  }
}
