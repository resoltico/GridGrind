package dev.erst.gridgrind.excel;

import org.apache.poi.poifs.crypt.HashAlgorithm;

/** Supported OOXML package-signature digest algorithms exposed by GridGrind. */
public enum ExcelOoxmlSignatureDigestAlgorithm {
  SHA256(HashAlgorithm.sha256),
  SHA384(HashAlgorithm.sha384),
  SHA512(HashAlgorithm.sha512);

  private final HashAlgorithm poiHashAlgorithm;

  ExcelOoxmlSignatureDigestAlgorithm(HashAlgorithm poiHashAlgorithm) {
    this.poiHashAlgorithm = poiHashAlgorithm;
  }

  HashAlgorithm poiHashAlgorithm() {
    return poiHashAlgorithm;
  }
}
