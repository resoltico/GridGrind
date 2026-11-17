package dev.erst.gridgrind.engine.media;

import java.io.IOException;
import java.util.Objects;

/** Synthetic valid media for integration tests whose subject is native workflow wiring. */
public final class MediaFixtures {
  private MediaFixtures() {}

  /** Returns independently encoded, complete PNG bytes. */
  public static byte[] validPng() throws IOException {
    try (var input =
        Objects.requireNonNull(MediaFixtures.class.getResourceAsStream("/media/valid.png"))) {
      return input.readAllBytes();
    }
  }
}
