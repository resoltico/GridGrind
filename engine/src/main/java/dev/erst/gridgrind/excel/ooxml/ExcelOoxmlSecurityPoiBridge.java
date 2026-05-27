package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.ExcelEnumMappingSupport;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.poifs.crypt.ChainingMode;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;

/** Maps OOXML encryption and signature enums between GridGrind and Apache POI. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelOoxmlSecurityPoiBridge {
  private static final Map<ExcelOoxmlEncryptionMode, EncryptionMode> TO_POI_ENCRYPTION_MODE =
      ExcelEnumMappingSupport.exactEnumMap(
          ExcelOoxmlEncryptionMode.class,
          "Apache POI OOXML encryption-mode mapping",
          Map.ofEntries(
              Map.entry(ExcelOoxmlEncryptionMode.AGILE, EncryptionMode.agile),
              Map.entry(ExcelOoxmlEncryptionMode.STANDARD, EncryptionMode.standard)));

  private static final Map<EncryptionMode, ExcelOoxmlEncryptionMode> FROM_POI_ENCRYPTION_MODE =
      Map.ofEntries(
          Map.entry(EncryptionMode.agile, ExcelOoxmlEncryptionMode.AGILE),
          Map.entry(EncryptionMode.standard, ExcelOoxmlEncryptionMode.STANDARD));

  private static final Map<CipherAlgorithm, ExcelOoxmlCipherAlgorithm> FROM_POI_CIPHER_ALGORITHM =
      ExcelEnumMappingSupport.exactEnumMap(
          CipherAlgorithm.class,
          "GridGrind OOXML cipher-algorithm mapping",
          Map.ofEntries(
              Map.entry(CipherAlgorithm.rc4, ExcelOoxmlCipherAlgorithm.RC4),
              Map.entry(CipherAlgorithm.aes128, ExcelOoxmlCipherAlgorithm.AES_128),
              Map.entry(CipherAlgorithm.aes192, ExcelOoxmlCipherAlgorithm.AES_192),
              Map.entry(CipherAlgorithm.aes256, ExcelOoxmlCipherAlgorithm.AES_256),
              Map.entry(CipherAlgorithm.rc2, ExcelOoxmlCipherAlgorithm.RC2),
              Map.entry(CipherAlgorithm.des, ExcelOoxmlCipherAlgorithm.DES),
              Map.entry(CipherAlgorithm.des3, ExcelOoxmlCipherAlgorithm.TRIPLE_DES),
              Map.entry(CipherAlgorithm.des3_112, ExcelOoxmlCipherAlgorithm.TRIPLE_DES_112),
              Map.entry(CipherAlgorithm.rsa, ExcelOoxmlCipherAlgorithm.RSA)));

  private static final Map<HashAlgorithm, ExcelOoxmlHashAlgorithm> FROM_POI_HASH_ALGORITHM =
      ExcelEnumMappingSupport.exactEnumMap(
          HashAlgorithm.class,
          "GridGrind OOXML hash-algorithm mapping",
          Map.ofEntries(
              Map.entry(HashAlgorithm.none, ExcelOoxmlHashAlgorithm.NONE),
              Map.entry(HashAlgorithm.sha1, ExcelOoxmlHashAlgorithm.SHA_1),
              Map.entry(HashAlgorithm.sha224, ExcelOoxmlHashAlgorithm.SHA_224),
              Map.entry(HashAlgorithm.sha256, ExcelOoxmlHashAlgorithm.SHA_256),
              Map.entry(HashAlgorithm.sha384, ExcelOoxmlHashAlgorithm.SHA_384),
              Map.entry(HashAlgorithm.sha512, ExcelOoxmlHashAlgorithm.SHA_512),
              Map.entry(HashAlgorithm.md2, ExcelOoxmlHashAlgorithm.MD2),
              Map.entry(HashAlgorithm.md4, ExcelOoxmlHashAlgorithm.MD4),
              Map.entry(HashAlgorithm.md5, ExcelOoxmlHashAlgorithm.MD5),
              Map.entry(HashAlgorithm.ripemd128, ExcelOoxmlHashAlgorithm.RIPEMD_128),
              Map.entry(HashAlgorithm.ripemd160, ExcelOoxmlHashAlgorithm.RIPEMD_160),
              Map.entry(HashAlgorithm.ripemd256, ExcelOoxmlHashAlgorithm.RIPEMD_256),
              Map.entry(HashAlgorithm.whirlpool, ExcelOoxmlHashAlgorithm.WHIRLPOOL)));

  private static final Map<ChainingMode, ExcelOoxmlChainingMode> FROM_POI_CHAINING_MODE =
      ExcelEnumMappingSupport.exactEnumMap(
          ChainingMode.class,
          "GridGrind OOXML chaining-mode mapping",
          Map.ofEntries(
              Map.entry(ChainingMode.ecb, ExcelOoxmlChainingMode.ECB),
              Map.entry(ChainingMode.cbc, ExcelOoxmlChainingMode.CBC),
              Map.entry(ChainingMode.cfb, ExcelOoxmlChainingMode.CFB)));

  private static final Map<ExcelOoxmlSignatureDigestAlgorithm, HashAlgorithm>
      TO_POI_SIGNATURE_DIGEST_ALGORITHM =
          ExcelEnumMappingSupport.exactEnumMap(
              ExcelOoxmlSignatureDigestAlgorithm.class,
              "Apache POI OOXML signature-digest mapping",
              Map.ofEntries(
                  Map.entry(ExcelOoxmlSignatureDigestAlgorithm.SHA256, HashAlgorithm.sha256),
                  Map.entry(ExcelOoxmlSignatureDigestAlgorithm.SHA384, HashAlgorithm.sha384),
                  Map.entry(ExcelOoxmlSignatureDigestAlgorithm.SHA512, HashAlgorithm.sha512)));

  private ExcelOoxmlSecurityPoiBridge() {}

  public static EncryptionMode toPoi(ExcelOoxmlEncryptionMode mode) {
    return ExcelEnumMappingSupport.requireMappedValue(
        TO_POI_ENCRYPTION_MODE, mode, "GridGrind OOXML encryption mode");
  }

  public static ExcelOoxmlEncryptionMode fromPoi(EncryptionMode mode) {
    return ExcelEnumMappingSupport.requireMappedValue(
        FROM_POI_ENCRYPTION_MODE, mode, "OOXML encryption mode");
  }

  public static ExcelOoxmlCipherAlgorithm fromPoi(CipherAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return ExcelEnumMappingSupport.requireMappedValue(
        FROM_POI_CIPHER_ALGORITHM, algorithm, "OOXML cipher algorithm");
  }

  public static ExcelOoxmlHashAlgorithm fromPoi(HashAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return ExcelEnumMappingSupport.requireMappedValue(
        FROM_POI_HASH_ALGORITHM, algorithm, "OOXML hash algorithm");
  }

  public static ExcelOoxmlChainingMode fromPoi(ChainingMode chainingMode) {
    Objects.requireNonNull(chainingMode, "chainingMode must not be null");
    return ExcelEnumMappingSupport.requireMappedValue(
        FROM_POI_CHAINING_MODE, chainingMode, "OOXML chaining mode");
  }

  public static HashAlgorithm toPoi(ExcelOoxmlSignatureDigestAlgorithm algorithm) {
    Objects.requireNonNull(algorithm, "algorithm must not be null");
    return ExcelEnumMappingSupport.requireMappedValue(
        TO_POI_SIGNATURE_DIGEST_ALGORITHM, algorithm, "GridGrind OOXML signature digest");
  }
}
