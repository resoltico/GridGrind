package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import java.util.Objects;
import org.apache.poi.poifs.crypt.ChainingMode;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
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

  static ExcelOoxmlCipherAlgorithm fromPoi(CipherAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return switch (algorithm) {
      case rc4 -> ExcelOoxmlCipherAlgorithm.RC4;
      case aes128 -> ExcelOoxmlCipherAlgorithm.AES_128;
      case aes192 -> ExcelOoxmlCipherAlgorithm.AES_192;
      case aes256 -> ExcelOoxmlCipherAlgorithm.AES_256;
      case rc2 -> ExcelOoxmlCipherAlgorithm.RC2;
      case des -> ExcelOoxmlCipherAlgorithm.DES;
      case des3 -> ExcelOoxmlCipherAlgorithm.TRIPLE_DES;
      case des3_112 -> ExcelOoxmlCipherAlgorithm.TRIPLE_DES_112;
      case rsa -> ExcelOoxmlCipherAlgorithm.RSA;
    };
  }

  static ExcelOoxmlHashAlgorithm fromPoi(HashAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return switch (algorithm) {
      case none -> ExcelOoxmlHashAlgorithm.NONE;
      case sha1 -> ExcelOoxmlHashAlgorithm.SHA_1;
      case sha224 -> ExcelOoxmlHashAlgorithm.SHA_224;
      case sha256 -> ExcelOoxmlHashAlgorithm.SHA_256;
      case sha384 -> ExcelOoxmlHashAlgorithm.SHA_384;
      case sha512 -> ExcelOoxmlHashAlgorithm.SHA_512;
      case md2 -> ExcelOoxmlHashAlgorithm.MD2;
      case md4 -> ExcelOoxmlHashAlgorithm.MD4;
      case md5 -> ExcelOoxmlHashAlgorithm.MD5;
      case ripemd128 -> ExcelOoxmlHashAlgorithm.RIPEMD_128;
      case ripemd160 -> ExcelOoxmlHashAlgorithm.RIPEMD_160;
      case ripemd256 -> ExcelOoxmlHashAlgorithm.RIPEMD_256;
      case whirlpool -> ExcelOoxmlHashAlgorithm.WHIRLPOOL;
    };
  }

  static ExcelOoxmlChainingMode fromPoi(ChainingMode chainingMode) {
    Objects.requireNonNull(chainingMode, "chainingMode must not be null");
    return switch (chainingMode) {
      case ecb -> ExcelOoxmlChainingMode.ECB;
      case cbc -> ExcelOoxmlChainingMode.CBC;
      case cfb -> ExcelOoxmlChainingMode.CFB;
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
