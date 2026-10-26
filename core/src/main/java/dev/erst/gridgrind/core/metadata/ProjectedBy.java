package dev.erst.gridgrind.core.metadata;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Read facets that make a factual record component available. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.RECORD_COMPONENT)
public @interface ProjectedBy {
  /** Facet identifiers that publish this component. */
  String[] value();
}
