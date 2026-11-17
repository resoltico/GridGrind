package dev.erst.gridgrind.engine.security;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies plaintext/replacement ownership transfer retains primary and secondary failures. */
class SecurityCleanupOwnershipTest {
  @TempDir Path root;

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // openAndRelease consumes the readable owner; an extra close would
  // repeat the injected failure.
  void failedPlaintextReleaseClosesTheAlreadyOpenedReplacementBeforePropagating() {
    var plaintextFailure = new IOException("synthetic plaintext deletion failure");
    var nativeFailure = new IOException("synthetic native close failure");
    var closes = new AtomicInteger();
    var readable =
        new SecurityAdapter.ReadableWorkbook(
            root.resolve("plain.xlsx"),
            OoxmlPackageSecurityReport.none(),
            Optional.empty(),
            () -> {
              throw plaintextFailure;
            });
    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                SecurityAdapter.openAndRelease(
                    readable,
                    ignored ->
                        (AutoCloseable)
                            () -> {
                              closes.incrementAndGet();
                              throw nativeFailure;
                            }));
    assertSame(plaintextFailure, failure);
    assertEquals(1, closes.get());
    assertEquals(1, failure.getSuppressed().length);
    assertSame(nativeFailure, failure.getSuppressed()[0]);
  }

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // The production ownership helper closes this readable after the
  // injected opener fails.
  void failedNativeOpenRetainsItsPrimaryFailureWhenPlaintextCleanupAlsoFails() {
    var opening = new IOException("synthetic open failure");
    var cleanup = new IOException("synthetic plaintext deletion failure");
    var readable =
        new SecurityAdapter.ReadableWorkbook(
            root.resolve("plain.xlsx"),
            OoxmlPackageSecurityReport.none(),
            Optional.empty(),
            () -> {
              throw cleanup;
            });
    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                SecurityAdapter.openAndRelease(
                    readable,
                    ignored -> {
                      throw opening;
                    }));
    assertSame(opening, failure);
    assertEquals(1, failure.getSuppressed().length);
    assertSame(cleanup, failure.getSuppressed()[0]);
  }
}
