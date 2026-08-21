package dev.erst.gridgrind.contract.dto;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Declares one request-record field's protocol presence and sensitivity semantics. */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.RECORD_COMPONENT)
public @interface ProtocolField {
  /** Whether the field may be omitted from an authored request. */
  boolean optional() default false;

  /** The effective default when an optional boolean field is omitted from the wire request. */
  ProtocolBooleanDefault booleanDefault() default ProtocolBooleanDefault.UNSPECIFIED;

  /** Whether diagnostics and telemetry must never reproduce the field value. */
  boolean secret() default false;
}
