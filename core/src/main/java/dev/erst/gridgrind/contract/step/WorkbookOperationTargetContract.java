package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;
import java.util.List;

/** Targeting behavior owned by one operation contract. */
sealed interface WorkbookOperationTargetContract
    permits StaticWorkbookOperationTargetContract, DerivedWorkbookOperationTargetContract {
  /** Returns selector classes accepted by this contract for the supplied bound operation. */
  List<Class<? extends Selector>> acceptedSelectors(Object operation);

  /** Returns the catalog surface published for this contract's targeting rule. */
  WorkbookStepTargeting.TargetSurface targetSurface();
}
