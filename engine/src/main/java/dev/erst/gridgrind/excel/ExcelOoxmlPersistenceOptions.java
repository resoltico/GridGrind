package dev.erst.gridgrind.excel;

/** Optional OOXML package-security settings applied during workbook persistence. */
public record ExcelOoxmlPersistenceOptions(
    ExcelOoxmlEncryptionOptions encryption, ExcelOoxmlSignatureOptions signature) {
  /** Returns whether neither encryption nor signing was requested for persistence. */
  public boolean isEmpty() {
    return encryption == null && signature == null;
  }
}
