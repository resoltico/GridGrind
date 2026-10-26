package dev.erst.gridgrind.core.vocabulary;

/** A validated domain value whose wire representation is one JSON string. */
public sealed interface WireString
    permits SheetName,
        SheetLabel,
        CellAddress,
        RangeRef,
        RangeLabel,
        DefinedName,
        DefinedNameLabel,
        TableName,
        RgbHex,
        StepId,
        PlanId {
  /** Identifies concrete validated records whose wire form is a scalar string. */
  static boolean isValueType(Class<?> type) {
    return type.isRecord() && WireString.class.isAssignableFrom(type);
  }

  /** Returns the validated text represented by this value. */
  String value();
}
