package dev.erst.gridgrind.cli.io;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.Objects;

/** Shared payload-output helper that normalizes one trailing newline without closing streams. */
final class CliPayloadOutput {
  private CliPayloadOutput() {}

  static void write(OutputStream outputStream, byte[] payload) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(payload, "payload must not be null");
    outputStream.write(payload);
    if (payload.length == 0 || payload[payload.length - 1] != '\n') {
      outputStream.write('\n');
    }
    outputStream.flush();
    // System.out and System.err swallow IOExceptions and retain only this failure flag.
    if (outputStream instanceof PrintStream console && console.checkError()) {
      throw new IOException("CLI output stream did not accept the payload.");
    }
  }
}
