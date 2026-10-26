package dev.erst.gridgrind.core.metadata;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Colocated catalog identity and documentation for a non-polymorphic record. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface CatalogType {
  /** Stable standalone catalog identity. */
  String id();

  /** Human-readable description of the record contract. */
  String summary();

  /** Shared catalog notes referenced by this record. */
  String[] noteRefs() default {};

  /** Presentation position among standalone record groups. */
  int displayOrder() default Integer.MAX_VALUE;
}
