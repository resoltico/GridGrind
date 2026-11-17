package dev.erst.gridgrind.core.source;

import java.io.IOException;
import java.io.InputStream;
import java.util.Base64;

/** Bounded-memory basic Base64 syntax and length validation for immutable source strings. */
final class Base64SourceValidation {
  private Base64SourceValidation() {}

  static void validate(String value) throws IOException {
    var characters = new Characters(value);
    try (InputStream decoder = Base64.getDecoder().wrap(characters)) {
      byte[] work = new byte[8192];
      while (decoder.read(work) != -1) {
        /* Discard decoded syntax-validation bytes. */
      }
      if (characters.position != value.length()) throw new IOException("Trailing base64 input");
    }
  }

  static long decodedSize(String value) {
    int length = value.length();
    long result = (long) (length / 4) * 3;
    int tail = length % 4;
    if (tail == 2) return result + 1;
    if (tail == 3) return result + 2;
    if (value.charAt(length - 1) == '=') result--;
    if (value.charAt(length - 2) == '=') result--;
    return result;
  }

  /** Read-only ASCII view that retains the exact source cursor without copying encoded content. */
  private static final class Characters extends InputStream {
    private final String value;
    private int position;

    private Characters(String value) {
      this.value = value;
    }

    @Override
    public int read() {
      if (position == value.length()) return -1;
      char next = value.charAt(position);
      position++;
      return next < 128 ? next : 255;
    }

    @Override
    public int read(byte[] target, int offset, int length) {
      java.util.Objects.checkFromIndexSize(offset, length, target.length);
      if (length == 0) return 0;
      int next = read();
      if (next < 0) return -1;
      target[offset] = (byte) next;
      return 1;
    }
  }
}
