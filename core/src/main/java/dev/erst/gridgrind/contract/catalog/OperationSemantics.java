package dev.erst.gridgrind.contract.catalog;

import java.util.ArrayList;
import java.util.EnumSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Canonical static effects, footprint, and preconditions for one concrete workbook operation. */
public record OperationSemantics(
    List<OperationEffect> effects,
    OperationEffectFootprint footprint,
    List<OperationPrecondition> preconditions) {
  public OperationSemantics {
    Objects.requireNonNull(effects, "effects must not be null");
    Set<OperationEffect> uniqueEffects = EnumSet.noneOf(OperationEffect.class);
    for (OperationEffect effect : effects) {
      uniqueEffects.add(Objects.requireNonNull(effect, "effects must not contain null values"));
    }
    if (uniqueEffects.isEmpty()) {
      throw new IllegalArgumentException("effects must not be empty");
    }
    effects = List.copyOf(new ArrayList<>(uniqueEffects));
    Objects.requireNonNull(footprint, "footprint must not be null");
    preconditions =
        List.copyOf(Objects.requireNonNull(preconditions, "preconditions must not be null"));
    for (OperationPrecondition precondition : preconditions) {
      Objects.requireNonNull(precondition, "preconditions must not contain null values");
    }
  }
}
