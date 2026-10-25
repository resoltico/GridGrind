package dev.erst.gridgrind.core.vocabulary;

/** A validated domain value whose wire representation is one JSON string. */
public sealed interface WireString
    permits SheetName, CellAddress, RangeRef, DefinedName, TableName, RgbHex, StepId, PlanId {
  /** Returns the validated text represented by this value. */
  String value();
}
