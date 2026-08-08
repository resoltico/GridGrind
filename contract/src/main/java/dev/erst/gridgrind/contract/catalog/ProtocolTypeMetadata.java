package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.selector.Selector;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Colocated public-contract metadata for one sealed protocol leaf type. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ProtocolTypeMetadata {
  /** Stable SCREAMING_SNAKE_CASE discriminator id for the protocol leaf. */
  String id();

  /** Public catalog/help summary for the protocol leaf. */
  String summary();

  /** Static selector families accepted by the leaf when targetingMode is STATIC. */
  Class<? extends Selector>[] targetSelectors() default {};

  /** Human-readable selector rule exposed to discovery/help surfaces. */
  String targetSelectorRule() default "";

  /** Strategy used to resolve target selectors for this leaf. */
  ProtocolTargetingMode targetingMode() default ProtocolTargetingMode.STATIC;
}
