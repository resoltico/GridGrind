package dev.erst.gridgrind.core.vocabulary;

/** A validated domain record represented by one primitive JSON value. */
public sealed interface WireScalar permits WireString, ProtocolVersion {
  /** Identifies concrete records whose component is their entire wire representation. */
  static boolean isValueType(Class<?> type) {
    return type.isRecord() && WireScalar.class.isAssignableFrom(type);
  }

  /** Returns the primitive string or integer represented by this value. */
  Object scalarValue();
}
