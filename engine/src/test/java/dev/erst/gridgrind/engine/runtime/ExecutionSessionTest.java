package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies run-owned resources unwind without erasing publication facts or host-owned data. */
class ExecutionSessionTest {
  @TempDir Path hostRoot;

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // Explicit close assertions verify failure order and that a second close
  // does not release callbacks twice.
  void everyResourceUnwindsAndPlaintextIsRemovedAfterIoAndRuntimeCloseFailures() throws Exception {
    Path marker = Files.writeString(hostRoot.resolve("host-marker"), "host-owned");
    var session = new ExecutionSession(hostRoot, ExecutionGrant.denied());
    Path plaintext = session.createTempFile("plaintext-", ".xlsx");
    Files.writeString(plaintext, "synthetic plaintext");
    Path ownedRoot = plaintext.getParent();
    var closed = new ArrayList<String>();
    var io = new IOException("synthetic descriptor close failure");
    var runtime = new IllegalStateException("synthetic callback close failure");
    session.own(() -> closed.add("first"));
    session.own(() -> {
      closed.add("second");
      throw io;
    });
    session.own(() -> {
      closed.add("third");
      throw runtime;
    });
    session.publication.returnedSuccessfully(hostRoot.resolve("output.xlsx").toString());
    PublicationOutcome before = session.publication.outcome();
    IOException failure = assertThrows(IOException.class, session::close);
    assertSame(runtime, failure.getCause());
    assertEquals(1, failure.getSuppressed().length);
    assertSame(io, failure.getSuppressed()[0]);
    assertEquals(List.of("third", "second", "first"), closed);
    assertFalse(Files.exists(plaintext));
    assertFalse(Files.exists(ownedRoot));
    assertEquals("host-owned", Files.readString(marker));
    assertTrue(Files.isDirectory(hostRoot));
    assertEquals(before, session.publication.outcome());
    assertEquals(
        PublicationOutcome.CommitObservation.RETURNED_SUCCESSFULLY,
        assertInstanceOf(PublicationOutcome.Uncertain.class, session.publication.outcome())
            .commitObservation());
    session.close();
    assertEquals(List.of("third", "second", "first"), closed);
  }

  @Test
  @SuppressWarnings(
      "PMD.CloseResource") // close is the fault-injection boundary; adding TWR would throw the
  // retained publication cleanup failure again.
  void publicationCleanupFailureRemainsPrimaryWithBoundedSecondaryErrorsAndNoSelfSuppression()
      throws Exception {
    var session = new ExecutionSession(hostRoot, ExecutionGrant.denied());
    Path plaintext = session.createTempFile("owned-", ".xlsx");
    Path ownedRoot = plaintext.getParent();
    var primary = new IOException("synthetic publication sibling cleanup failure");
    session.publication.cleanupFailed(primary);
    session.publication.cleanupFailed(primary);
    var attempts = new AtomicInteger();
    session.own(() -> {
      attempts.incrementAndGet();
      throw primary;
    });
    for (int index = 0; index < 12; index++) {
      session.own(() -> {
        attempts.incrementAndGet();
        throw new IOException("secondary cleanup failure");
      });
    }
    IOException failure = assertThrows(IOException.class, session::close);
    assertSame(primary, failure);
    assertEquals(13, attempts.get());
    assertEquals(8, failure.getSuppressed().length);
    for (Throwable secondary : failure.getSuppressed()) assertNotSame(primary, secondary);
    assertFalse(Files.exists(ownedRoot));
    assertTrue(Files.isDirectory(hostRoot));
  }

  @Test
  void finalRootCleanupDoesNotFollowAnOwnedSymlinkIntoHostFiles() throws Exception {
    Path hostData = Files.createDirectory(hostRoot.resolve("host-data"));
    Files.writeString(hostData.resolve("keep"), "host-owned");
    Path owned;
    try (var session = new ExecutionSession(hostRoot, ExecutionGrant.denied())) {
      owned = session.createTempFile("owned-", ".xlsx").getParent();
      Files.createSymbolicLink(owned.resolve("host-link"), hostData);
    }
    assertFalse(Files.exists(owned));
    assertEquals("host-owned", Files.readString(hostData.resolve("keep")));
  }
}
