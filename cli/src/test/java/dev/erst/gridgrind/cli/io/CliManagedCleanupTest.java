package dev.erst.gridgrind.cli.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Proxy;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SecureDirectoryStream;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises descriptor failure unwinding while preserving caller-owned files and roots. */
class CliManagedCleanupTest {
  @TempDir Path hostRoot;

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // The asserted managed close consumes the resolver; an extra close would
  // reinvoke the failure seam.
  void descriptorCloseFailureStillDeletesOwnedScratchAndPreservesHostFiles() throws Exception {
    Path hostData = Files.createDirectory(hostRoot.resolve("host-data"));
    Files.writeString(hostData.resolve("keep"), "host-owned");
    Path scratch = Files.createDirectory(hostRoot.resolve("owned"));
    Files.writeString(scratch.resolve("plaintext"), "owned plaintext");
    Files.createSymbolicLink(scratch.resolve("host-link"), hostData);
    var closes = new AtomicInteger();
    var marker = new IOException("synthetic provider close failure");
    try (var stream =
        Files.newDirectoryStream(Files.createDirectory(hostRoot.resolve("secrets")))) {
      var resolver = failingResolver(stream, marker, closes);
      var managed =
          new CliExecutionBindingsFactory.ManagedRequestInputs(
              new GridGrindRequestInputs(hostRoot, scratch, ExecutionGrant.denied()),
              Optional.of(resolver));
      assertSame(marker, assertThrows(IOException.class, managed::close));
      assertEquals(1, closes.get());
      assertFalse(Files.exists(scratch));
      assertEquals("host-owned", Files.readString(hostData.resolve("keep")));
    }
  }

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // The asserted managed close consumes the resolver while the native
  // stream has its own TWR fallback.
  void providerErrorPropagatesAfterOwnedScratchHasUnwound() throws Exception {
    Path scratch = Files.createDirectory(hostRoot.resolve("owned"));
    Files.writeString(scratch.resolve("plaintext"), "owned plaintext");
    var closes = new AtomicInteger();
    var marker = new AssertionError("synthetic provider close Error");
    try (var stream =
        Files.newDirectoryStream(Files.createDirectory(hostRoot.resolve("secrets")))) {
      var resolver = failingResolver(stream, marker, closes);
      var managed =
          new CliExecutionBindingsFactory.ManagedRequestInputs(
              new GridGrindRequestInputs(hostRoot, scratch, ExecutionGrant.denied()),
              Optional.of(resolver));
      assertSame(marker, assertThrows(AssertionError.class, managed::close));
      assertEquals(1, closes.get());
      assertFalse(Files.exists(scratch));
    }
  }

  @SuppressWarnings("unchecked") // The proxy delegates the same Path-typed directory capability.
  private static CliFileSecretResolver failingResolver(
      java.nio.file.DirectoryStream<Path> stream, Throwable failure, AtomicInteger closes)
      throws IOException {
    if (!(stream instanceof SecureDirectoryStream<Path> secure)) {
      throw new IOException("Test requires the qualified secure directory provider");
    }
    var wrapper =
        (SecureDirectoryStream<Path>)
            Proxy.newProxyInstance(
                Thread.currentThread().getContextClassLoader(),
                new Class<?>[] {SecureDirectoryStream.class},
                (proxy, method, arguments) -> {
                  if ("close".equals(method.getName())) {
                    closes.incrementAndGet();
                    secure.close();
                    throw failure;
                  }
                  try {
                    return method.invoke(secure, arguments);
                  } catch (InvocationTargetException exception) {
                    throw exception.getCause();
                  }
                });
    return new CliFileSecretResolver(wrapper);
  }
}
