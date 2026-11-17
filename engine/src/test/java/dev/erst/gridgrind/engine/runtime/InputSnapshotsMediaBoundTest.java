package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.engine.media.MediaIssues;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;

/** A media caller's smaller cap governs acquisition before shared snapshot buffers grow. */
class InputSnapshotsMediaBoundTest {
  @Test
  void oversizedAcquisitionReadsOnlyTheRemainingAllowanceAndOneByteThenPoisons() throws Exception {
    var snapshots = new InputSnapshots(grant());
    var read = new AtomicInteger();
    var closed = new AtomicInteger();
    var refusal = MediaIssues.limit(Limit.ID_AUTHORED_MEDIA_BYTES);
    var bound = new InputSnapshotLimit(3, refusal);
    dev.erst.gridgrind.engine.api.GridGrindPayloadOpener source =
        () ->
            new java.io.InputStream() {
              private final ByteArrayInputStream data = new ByteArrayInputStream(new byte[100]);

              @Override
              public int read() {
                int next = data.read();
                if (next >= 0) read.incrementAndGet();
                return next;
              }

              @Override
              public int read(byte[] target, int offset, int length) {
                int count = data.read(target, offset, length);
                read.addAndGet(Math.max(0, count));
                return count;
              }

              @Override
              public void close() throws java.io.IOException {
                closed.incrementAndGet();
                data.close();
              }
            };
    assertSame(
        refusal,
        assertThrows(RuntimeException.class, () -> snapshots.standardInputWithin(source, bound)));
    assertEquals(4, read.get());
    assertEquals(1, closed.get());
    assertSame(
        refusal,
        assertThrows(
            RuntimeException.class,
            () ->
                snapshots.standardInput(() -> {
                  throw new AssertionError("Poisoned snapshot was reopened");
                })));
  }

  @Test
  void copyChecksCapBeforeWritingOverflowAndPreservesTheExactLimit() throws Exception {
    var exact = new ByteArrayOutputStream();
    var bound = new InputSnapshotLimit(3, MediaIssues.limit(Limit.ID_AUTHORED_MEDIA_BYTES));
    new InputSnapshots(grant())
        .copyWithin(new ByteArrayInputStream(new byte[] {1, 2, 3}), exact, bound);
    assertEquals(3, exact.size());
    var overflow = new ByteArrayOutputStream();
    assertThrows(
        RuntimeException.class,
        () ->
            new InputSnapshots(grant())
                .copyWithin(new ByteArrayInputStream(new byte[4]), overflow, bound));
    assertEquals(0, overflow.size());
  }

  @Test
  void completedLargerStdinSnapshotSurvivesASmallerMediaCallerRefusal() throws Exception {
    var snapshots = new InputSnapshots(grant());
    var opens = new AtomicInteger();
    var opener =
        (dev.erst.gridgrind.engine.api.GridGrindPayloadOpener)
            () -> {
              opens.incrementAndGet();
              return new ByteArrayInputStream(new byte[4]);
            };
    assertEquals(4, snapshots.standardInput(opener).remaining());
    var refusal = MediaIssues.limit(Limit.ID_AUTHORED_MEDIA_BYTES);
    assertSame(
        refusal,
        assertThrows(
            RuntimeException.class,
            () -> snapshots.standardInputWithin(opener, new InputSnapshotLimit(3, refusal))));
    assertEquals(4, snapshots.standardInput(opener).remaining());
    assertEquals(1, opens.get());
  }

  private static ExecutionGrant grant() {
    return new ExecutionGrant(
        1_000,
        List.of(new ExecutionGrant.ReadAuthority.Stdin()),
        List.of(),
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new ExecutionGrant.PublicationAuthority.None(),
        List.of(),
        AcceptancePolicy.minimum());
  }
}
