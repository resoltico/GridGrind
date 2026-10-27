package dev.erst.gridgrind.fuzz;

import java.io.IOException;

/** Classic standalone-driver entry point; ordinary seed tests never load the native driver. */
public final class FuzzTarget {
  private static final Harness HARNESS =
      Harness.fromKey(System.getProperty("gridgrind.fuzz.harness", ""));

  private FuzzTarget() {}

  public static void fuzzerTestOneInput(byte[] input) throws IOException {
    HARNESS.check(input);
  }
}
