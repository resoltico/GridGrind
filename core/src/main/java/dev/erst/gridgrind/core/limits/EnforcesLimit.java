package dev.erst.gridgrind.core.limits;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Connects a concrete enforcement owner to its canonical limit entries for generated references.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface EnforcesLimit {
  /** Stable IDs declared by the limit catalog. */
  String[] value();
}
