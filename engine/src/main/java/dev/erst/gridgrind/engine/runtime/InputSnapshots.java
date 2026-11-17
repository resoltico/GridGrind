package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindPayloadOpener;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** One invocation's charged input reads and immutable standard-input snapshot. */
final class InputSnapshots {
  private final ExecutionGrant grant;
  private long consumed;
  private java.util.Optional<Exception> failure = java.util.Optional.empty();
  private @Nullable ByteBuffer standardInput;

  InputSnapshots(ExecutionGrant grant) {
    this.grant = Objects.requireNonNull(grant);
  }

  ByteBuffer standardInput(GridGrindPayloadOpener opener) throws IOException {
    return captureStandardInput(opener, java.util.Optional.empty());
  }

  ByteBuffer standardInputWithin(GridGrindPayloadOpener opener, InputSnapshotLimit limit)
      throws IOException {
    return captureStandardInput(opener, java.util.Optional.of(limit));
  }

  private ByteBuffer captureStandardInput(
      GridGrindPayloadOpener opener, java.util.Optional<InputSnapshotLimit> limit)
      throws IOException {
    var cached = standardInput;
    if (cached != null) {
      limit.ifPresent(bound -> bound.require(cached.remaining()));
      return cached.asReadOnlyBuffer();
    }
    requireAcquisitionAvailable();
    try (InputStream input =
        Objects.requireNonNull(opener.open(), "payload opener returned null")) {
      var output = new ByteArrayOutputStream();
      copyWithLimit(input, output, limit);
      standardInput = ByteBuffer.wrap(output.toByteArray()).asReadOnlyBuffer();
    } catch (IOException | RuntimeException exception) {
      failure = java.util.Optional.of(exception);
      throw exception;
    }
    return standardInput.asReadOnlyBuffer();
  }

  void requireAcquisitionAvailable() throws IOException {
    if (failure.isPresent()) {
      Exception previous = failure.orElseThrow();
      if (previous instanceof IOException io) throw io;
      throw (RuntimeException) previous;
    }
  }

  void copy(InputStream input, OutputStream output) throws IOException {
    copyWithLimit(input, output, java.util.Optional.empty());
  }

  void copyWithin(InputStream input, OutputStream output, InputSnapshotLimit limit)
      throws IOException {
    copyWithLimit(input, output, java.util.Optional.of(limit));
  }

  private void copyWithLimit(
      InputStream input, OutputStream output, java.util.Optional<InputSnapshotLimit> limit)
      throws IOException {
    requireAcquisitionAvailable();
    try {
      copyAvailable(input, output, limit);
    } catch (IOException | RuntimeException exception) {
      failure = java.util.Optional.of(exception);
      throw exception;
    }
  }

  private void copyAvailable(
      InputStream input, OutputStream output, java.util.Optional<InputSnapshotLimit> limit)
      throws IOException {
    byte[] buffer = new byte[8192];
    long resourceBytes = 0;
    while (true) {
      long remaining = grant.maxInputSnapshotBytes() - consumed;
      if (limit.isPresent())
        remaining = Math.min(remaining, limit.orElseThrow().maxBytes() - resourceBytes);
      int length =
          (int) Math.min(buffer.length, remaining == Long.MAX_VALUE ? remaining : remaining + 1);
      int read = input.read(buffer, 0, length);
      if (read < 0) return;
      if (read == 0) {
        int next = input.read();
        if (next < 0) return;
        charge(1);
        resourceBytes++;
        long count = resourceBytes;
        limit.ifPresent(bound -> bound.require(count));
        output.write(next);
      } else {
        charge(read);
        resourceBytes += read;
        long count = resourceBytes;
        limit.ifPresent(bound -> bound.require(count));
        output.write(buffer, 0, read);
      }
    }
  }

  private void charge(int bytes) {
    long previous = consumed;
    consumed += Math.min(bytes, grant.maxInputSnapshotBytes() - consumed);
    grant.requireInputSnapshotBytes(previous, bytes);
  }
}
