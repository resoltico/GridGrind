package dev.erst.gridgrind.excel;

/** Signature-validation state for one OOXML package signature part. */
public enum ExcelOoxmlSignatureState {
  VALID,
  INVALID,
  INVALIDATED_BY_MUTATION;

  ExcelOoxmlSignatureState afterMutation() {
    return this == VALID ? INVALIDATED_BY_MUTATION : this;
  }
}
