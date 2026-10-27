package dev.erst.gridgrind.core.metadata;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Container for scalar constraints on one component. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.RECORD_COMPONENT)
public @interface FieldRules {
  /** All scalar rules declared for this component. */
  FieldRule[] value();
}
