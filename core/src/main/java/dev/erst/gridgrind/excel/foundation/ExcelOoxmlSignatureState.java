package dev.erst.gridgrind.excel.foundation;

/** Signature-validation state for one OOXML package signature part. */
public enum ExcelOoxmlSignatureState {
  VALID,
  INVALID,
  INVALIDATED_BY_MUTATION;

  /** Returns the persisted state after a workbook mutation invalidates live signatures. */
  public ExcelOoxmlSignatureState afterMutation() {
    return this == VALID ? INVALIDATED_BY_MUTATION : this;
  }
}
