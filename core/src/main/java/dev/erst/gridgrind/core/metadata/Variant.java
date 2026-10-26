package dev.erst.gridgrind.core.metadata;

import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.semantics.ProtocolTargetingMode;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Colocated public-contract metadata for one sealed protocol leaf type. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Variant {
  /** Stable SCREAMING_SNAKE_CASE discriminator id for the protocol leaf. */
  String id();

  /** Public catalog/help summary for the protocol leaf. */
  String summary();

  /** Shared static effects and footprint selected for this operation. */
  dev.erst.gridgrind.core.semantics.OperationProfile profile() default
      dev.erst.gridgrind.core.semantics.OperationProfile.NONE;

  /** Precondition records required by this operation. */
  Class<? extends dev.erst.gridgrind.core.semantics.OperationPrecondition>[]
      preconditions() default {};

  /** Shared catalog note identifiers attached to this variant. */
  String[] noteRefs() default {};

  /** Static selector families accepted by the leaf when targetingMode is STATIC. */
  Class<? extends Selector>[] targetSelectors() default {};

  /** Human-readable selector rule exposed to discovery/help surfaces. */
  String targetSelectorRule() default "";

  /** Strategy used to resolve target selectors for this leaf. */
  ProtocolTargetingMode targetingMode() default ProtocolTargetingMode.STATIC;
}
