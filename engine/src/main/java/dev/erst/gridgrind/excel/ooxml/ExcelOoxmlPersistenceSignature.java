package dev.erst.gridgrind.excel.ooxml;

import java.util.Objects;

/** Explicit package-signature disposition for one persisted OOXML package. */
public sealed interface ExcelOoxmlPersistenceSignature
    permits ExcelOoxmlPersistenceSignature.Unsigned, ExcelOoxmlPersistenceSignature.Sign {
  /** Deliberately persist without an OOXML package signature. */
  record Unsigned() implements ExcelOoxmlPersistenceSignature {}

  /** Persist with one explicit OOXML package signature. */
  record Sign(ExcelOoxmlSignatureOptions options) implements ExcelOoxmlPersistenceSignature {
    public Sign {
      Objects.requireNonNull(options, "options must not be null");
    }
  }
}
