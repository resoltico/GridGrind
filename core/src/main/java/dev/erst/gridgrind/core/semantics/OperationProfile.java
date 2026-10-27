package dev.erst.gridgrind.core.semantics;

import java.util.List;

/** Shared conservative effect declarations selected explicitly by operation metadata. */
@SuppressWarnings("ImmutableEnumChecker") // Each effect list is frozen by List.copyOf.
public enum OperationProfile {
  NONE(List.of(), OperationEffectFootprint.TARGET_BOUNDED),
  READ(List.of(OperationEffect.READ_WORKBOOK), OperationEffectFootprint.TARGET_BOUNDED),
  CELL_MUTATION(
      List.of(
          OperationEffect.READ_WORKBOOK,
          OperationEffect.MUTATE_WORKBOOK,
          OperationEffect.CELLS,
          OperationEffect.FORMULAS,
          OperationEffect.STYLES,
          OperationEffect.HYPERLINKS,
          OperationEffect.COMMENTS,
          OperationEffect.DRAWINGS),
      OperationEffectFootprint.CONSERVATIVE_WORKBOOK_WIDE),
  DRAWING_MUTATION(
      List.of(
          OperationEffect.READ_WORKBOOK,
          OperationEffect.MUTATE_WORKBOOK,
          OperationEffect.CELLS,
          OperationEffect.FORMULAS,
          OperationEffect.DRAWINGS),
      OperationEffectFootprint.CONSERVATIVE_WORKBOOK_WIDE),
  STRUCTURED_MUTATION(
      List.of(
          OperationEffect.READ_WORKBOOK,
          OperationEffect.MUTATE_WORKBOOK,
          OperationEffect.CELLS,
          OperationEffect.FORMULAS,
          OperationEffect.TABLES,
          OperationEffect.NAMED_RANGES,
          OperationEffect.DATA_VALIDATIONS,
          OperationEffect.CONDITIONAL_FORMATTING,
          OperationEffect.AUTOFILTERS,
          OperationEffect.PIVOT_TABLES,
          OperationEffect.CUSTOM_XML),
      OperationEffectFootprint.CONSERVATIVE_WORKBOOK_WIDE),
  WORKBOOK_MUTATION(
      List.of(
          OperationEffect.READ_WORKBOOK,
          OperationEffect.MUTATE_WORKBOOK,
          OperationEffect.CELLS,
          OperationEffect.FORMULAS,
          OperationEffect.STYLES,
          OperationEffect.HYPERLINKS,
          OperationEffect.COMMENTS,
          OperationEffect.DRAWINGS,
          OperationEffect.WORKBOOK_STRUCTURE,
          OperationEffect.SHEET_STRUCTURE,
          OperationEffect.ROWS,
          OperationEffect.COLUMNS,
          OperationEffect.TABLES,
          OperationEffect.NAMED_RANGES,
          OperationEffect.DATA_VALIDATIONS,
          OperationEffect.CONDITIONAL_FORMATTING,
          OperationEffect.AUTOFILTERS,
          OperationEffect.PIVOT_TABLES,
          OperationEffect.WORKBOOK_METADATA,
          OperationEffect.WORKBOOK_SECURITY),
      OperationEffectFootprint.CONSERVATIVE_WORKBOOK_WIDE);

  private final List<OperationEffect> effects;
  private final OperationEffectFootprint footprint;

  OperationProfile(List<OperationEffect> effects, OperationEffectFootprint footprint) {
    this.effects = List.copyOf(effects);
    this.footprint = footprint;
  }

  /** Returns the immutable conservative effect set. */
  public List<OperationEffect> effects() {
    return effects;
  }

  /** Returns the declared scope of those effects. */
  public OperationEffectFootprint footprint() {
    return footprint;
  }
}
