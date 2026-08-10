package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Selector families derived from a bound nested operation. */
record DerivedWorkbookOperationTargetContract(
    String rule, WorkbookOperationTargetSelectorDerivation derivation)
    implements WorkbookOperationTargetContract {
  DerivedWorkbookOperationTargetContract {
    if (rule == null || rule.isBlank()) {
      throw new IllegalArgumentException("rule must not be blank");
    }
    Objects.requireNonNull(derivation, "derivation must not be null");
  }

  @Override
  public List<Class<? extends Selector>> acceptedSelectors(Object operation) {
    Objects.requireNonNull(operation, "operation must not be null");
    return List.of(derivation.acceptedSelectors(operation));
  }

  @Override
  public WorkbookStepTargeting.TargetSurface targetSurface() {
    return new WorkbookStepTargeting.TargetSurface(List.of(), Optional.of(rule));
  }
}
