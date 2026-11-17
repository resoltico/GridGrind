package dev.erst.gridgrind.cli.io;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.source.SecretReference;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Mutable provider sizes cannot bypass actual-read limits or silently truncate secrets. */
class CliSecretBoundedReadTest {
  @TempDir Path root;

  @ParameterizedTest
  @ValueSource(booleans = {false, true})
  void growthAfterSizeHintOrDuringReadsRejectsWithoutExposingSecret(boolean duringRead)
      throws Exception {
    Path leaf = Files.writeString(root.resolve("synthetic-password"), "x");
    var closed = new AtomicBoolean();
    var count = new AtomicInteger();
    try (DirectoryStream<Path> stream = Files.newDirectoryStream(root);
        CliFileSecretResolver resolver =
            new CliFileSecretResolver(
                new FaultInjectingSecureDirectoryStream(
                    (SecureDirectoryStream<Path>) stream,
                    channel ->
                        new ObservedChannel(channel, closed, count) {
                          private boolean first = true;

                          @Override
                          public long size() throws IOException {
                            long original = super.size();
                            if (!duringRead) Files.writeString(leaf, "x".repeat(65_537));
                            return original;
                          }

                          @Override
                          public int read(ByteBuffer destination) throws IOException {
                            if (duringRead && !first) Files.writeString(leaf, "x".repeat(65_537));
                            int limit = destination.limit();
                            if (duringRead && first) destination.limit(destination.position() + 1);
                            first = false;
                            try {
                              return super.read(destination);
                            } finally {
                              destination.limit(limit);
                            }
                          }
                        },
                    false))) {
      var failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> resolver.resolve(new SecretReference("synthetic-password")));
      assertEquals("secret provider file exceeds 65536 bytes", failure.getMessage());
      assertEquals(65_537, count.get(), "At most one overflow lookahead byte");
      assertTrue(closed.get(), "Rejected leaf descriptor must close before returning");
    }
  }

  @Test
  void exactLimitReadsToEofAndMalformedUtf8ClosesTheDescriptor() throws Exception {
    Path leaf = Files.writeString(root.resolve("synthetic-password"), "x".repeat(65_536));
    var closed = new AtomicBoolean();
    var count = new AtomicInteger();
    try (DirectoryStream<Path> stream = Files.newDirectoryStream(root);
        CliFileSecretResolver resolver =
            new CliFileSecretResolver(
                new FaultInjectingSecureDirectoryStream(
                    (SecureDirectoryStream<Path>) stream,
                    channel -> new ObservedChannel(channel, closed, count),
                    false))) {
      char[] secret = resolver.resolve(new SecretReference("synthetic-password"));
      try {
        assertEquals(65_536, secret.length);
      } finally {
        java.util.Arrays.fill(secret, '\0');
      }
      assertEquals(65_536, count.get());
      assertTrue(closed.get());
      Files.write(leaf, new byte[] {(byte) 0xC3, (byte) 0x28});
      closed.set(false);
      var failure =
          assertThrows(
              IllegalArgumentException.class,
              () -> resolver.resolve(new SecretReference("synthetic-password")));
      assertEquals("secret provider file is not valid UTF-8", failure.getMessage());
      assertTrue(closed.get());
    }
  }

  /** Counts actual provider bytes and observes transferred descriptor closure. */
  private static class ObservedChannel implements SeekableByteChannel {
    private final SeekableByteChannel delegate;
    private final AtomicBoolean closed;
    private final AtomicInteger count;

    ObservedChannel(SeekableByteChannel delegate, AtomicBoolean closed, AtomicInteger count) {
      this.delegate = delegate;
      this.closed = closed;
      this.count = count;
    }

    @Override
    public int read(ByteBuffer bytes) throws IOException {
      int result = delegate.read(bytes);
      if (result > 0) count.addAndGet(result);
      return result;
    }

    @Override
    public long size() throws IOException {
      return delegate.size();
    }

    @Override
    public boolean isOpen() {
      return delegate.isOpen();
    }

    @Override
    public void close() throws IOException {
      delegate.close();
      closed.set(true);
    }

    @Override
    public long position() throws IOException {
      return delegate.position();
    }

    @Override
    public SeekableByteChannel position(long value) throws IOException {
      delegate.position(value);
      return this;
    }

    @Override
    public SeekableByteChannel truncate(long value) throws IOException {
      delegate.truncate(value);
      return this;
    }

    @Override
    public int write(ByteBuffer bytes) throws IOException {
      return delegate.write(bytes);
    }
  }
}
