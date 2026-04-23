package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import java.util.Objects;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;

/** Maps OOXML encryption and signature enums between GridGrind and Apache POI. */
final class ExcelOoxmlSecurityPoiBridge {
  private ExcelOoxmlSecurityPoiBridge() {}

  static EncryptionMode toPoi(ExcelOoxmlEncryptionMode mode) {
    return switch (mode) {
      case AGILE -> EncryptionMode.agile;
      case STANDARD -> EncryptionMode.standard;
    };
  }

  static ExcelOoxmlEncryptionMode fromPoi(EncryptionMode mode) {
    return switch (mode) {
      case agile -> ExcelOoxmlEncryptionMode.AGILE;
      case standard -> ExcelOoxmlEncryptionMode.STANDARD;
      default -> throw new IllegalArgumentException("Unsupported OOXML encryption mode: " + mode);
    };
  }

  static HashAlgorithm toPoi(ExcelOoxmlSignatureDigestAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return switch (algorithm) {
      case SHA256 -> HashAlgorithm.sha256;
      case SHA384 -> HashAlgorithm.sha384;
      case SHA512 -> HashAlgorithm.sha512;
    };
  }
}
