package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HexFormat;
import java.util.Objects;

/** Computes the immutable SHA-256 identity recorded for a private request input copy. */
final class RequestPathDigest {
  private RequestPathDigest() {}

  static String sha256(Path path) throws IOException {
    return sha256(path, MessageDigest::getInstance);
  }

  static String sha256(Path path, DigestFactory digestFactory) throws IOException {
    Objects.requireNonNull(path, "path must not be null");
    Objects.requireNonNull(digestFactory, "digestFactory must not be null");
    try (InputStream input = Files.newInputStream(path)) {
      MessageDigest digest = sha256Digest(digestFactory);
      byte[] buffer = new byte[16_384];
      while (true) {
        int read = input.read(buffer);
        if (read < 0) {
          break;
        }
        digest.update(buffer, 0, read);
      }
      return HexFormat.of().formatHex(digest.digest());
    }
  }

  private static MessageDigest sha256Digest(DigestFactory digestFactory) {
    try {
      return digestFactory.create("SHA-256");
    } catch (NoSuchAlgorithmException exception) {
      throw new IllegalStateException("SHA-256 must be available in the Java runtime", exception);
    }
  }

  /** Creates the required SHA-256 implementation for one materialized input identity. */
  @FunctionalInterface
  interface DigestFactory {
    /** Creates the requested digest implementation. */
    MessageDigest create(String algorithm) throws NoSuchAlgorithmException;
  }
}
