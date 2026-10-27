package dev.erst.gridgrind.core.metadata;

import java.lang.annotation.ElementType;
import java.lang.annotation.Repeatable;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** One scalar wire constraint, colocated with its record component. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.RECORD_COMPONENT)
@Repeatable(FieldRules.class)
public @interface FieldRule {
  /** Constraint category represented by this annotation. */
  Kind kind();

  /** Pattern or suffix text for text-valued constraints. */
  String text() default "";

  /** Inclusive lower bound for length or numeric constraints. */
  double minimum() default 0;

  /** Inclusive upper bound for length or numeric constraints. */
  double maximum() default 0;

  /** Supported scalar constraint categories. */
  enum Kind {
    NON_BLANK,
    STRING_PATTERN,
    LENGTH_RANGE,
    NUMBER_RANGE,
    INTEGRAL,
    PATH_SUFFIX
  }
}
