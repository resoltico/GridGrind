package dev.erst.gridgrind.core.metadata;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Declares the wire default applied to an absent record component before binding. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.RECORD_COMPONENT)
public @interface WireDefault {
  /** Literal text, enum constant, boolean, or empty-list marker for an absent component. */
  String value() default "";

  /** Core factory owner, or the component type when omitted. */
  Class<?> factory() default Void.class;

  /** Static factory method, or an empty string for its public no-argument constructor. */
  String method() default "";

  /** Other component names supplied as factory arguments. */
  String[] arguments() default {};

  /** Whether the effective default is omitted during wire serialization. */
  boolean omit() default false;

  /** Sibling component copied when this component is absent. */
  String from() default "";
}
