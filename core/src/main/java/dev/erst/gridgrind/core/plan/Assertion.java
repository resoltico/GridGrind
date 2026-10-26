package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.semantics.WorkbookOperationContracts;
import java.util.Objects;

/** First-class verification contract evaluated against a workbook target. */
@Union(discriminator = "type")
public sealed interface Assertion
    permits PresenceAssertion,
        CellAssertion,
        WorkbookFactAssertion,
        AnalysisAssertion,
        CompositeAssertion {

  /** Stable SCREAMING_SNAKE_CASE discriminator mirrored in catalog and result surfaces. */
  default String assertionType() {
    return GridGrindProtocolTypeNames.assertionTypeName(getClass().asSubclass(Assertion.class));
  }

  /** Returns the selector types accepted by one assertion instance. */
  static Class<? extends Selector>[] allowedTargetTypes(Assertion assertion) {
    Objects.requireNonNull(assertion, "assertion must not be null");
    return WorkbookOperationContracts.targetSelectorsFor(assertion);
  }
}
