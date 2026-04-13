package dev.erst.gridgrind.excel;

import org.apache.poi.poifs.crypt.EncryptionMode;

/** Supported OOXML package-encryption families exposed by GridGrind. */
public enum ExcelOoxmlEncryptionMode {
  AGILE(EncryptionMode.agile),
  STANDARD(EncryptionMode.standard);

  private final EncryptionMode poiMode;

  ExcelOoxmlEncryptionMode(EncryptionMode poiMode) {
    this.poiMode = poiMode;
  }

  EncryptionMode poiMode() {
    return poiMode;
  }

  static ExcelOoxmlEncryptionMode fromPoi(EncryptionMode poiMode) {
    return switch (poiMode) {
      case agile -> AGILE;
      case standard -> STANDARD;
      default ->
          throw new IllegalArgumentException("Unsupported OOXML encryption mode: " + poiMode);
    };
  }
}
