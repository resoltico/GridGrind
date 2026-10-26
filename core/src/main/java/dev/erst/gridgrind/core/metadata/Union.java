package dev.erst.gridgrind.core.metadata;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Declares a sealed wire union and its discriminator field. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Union {
  /** Wire tag field, or empty for a structural envelope union. */
  String discriminator() default "type";

  /** Whether this union is an independently discoverable catalog group. */
  boolean catalog() default false;

  /** Presentation position among discoverable union groups. */
  int displayOrder() default Integer.MAX_VALUE;
}
