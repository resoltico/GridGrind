package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies authorized, bounded snapshot acquisition, immutable reuse and stream ownership. */
class InputSnapshotsTest {
  @TempDir Path root;

  private ExecutionGrant grant(long bytes, List<ExecutionGrant.ReadAuthority> reads) {
    return new ExecutionGrant(
        bytes,
        reads,
        List.of(),
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new ExecutionGrant.PublicationAuthority.None(),
        List.of(),
        AcceptancePolicy.minimum());
  }

  @Test
  void deniedPayloadOpenerIsNeverInvoked() {
    var opens = new AtomicInteger();
    var inputs =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            ExecutionGrant.denied(),
            Optional.of(() -> {
              opens.incrementAndGet();
              throw new AssertionError("Denied payload was opened");
            }),
            Optional.empty());
    assertThrows(ExecutionAuthorityDeniedException.class, inputs::standardInputSnapshot);
    assertEquals(0, opens.get());
  }

  @Test
  void exactBudgetSnapshotsStdinOnceAndClosesTransferredStream() throws Exception {
    var opens = new AtomicInteger();
    var closes = new AtomicInteger();
    var inputs =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            grant(3, List.of(new ExecutionGrant.ReadAuthority.Stdin())),
            Optional.of(() -> {
              opens.incrementAndGet();
              return new ByteArrayInputStream(new byte[] {1, 2, 3}) {
                @Override
                public void close() {
                  closes.incrementAndGet();
                }
              };
            }),
            Optional.empty());
    assertEquals(
        ByteBuffer.wrap(new byte[] {1, 2, 3}), inputs.standardInputSnapshot().orElseThrow());
    assertEquals(
        ByteBuffer.wrap(new byte[] {1, 2, 3}), inputs.standardInputSnapshot().orElseThrow());
    var firstView = inputs.standardInputSnapshot().orElseThrow();
    assertThrows(java.nio.ReadOnlyBufferException.class, () -> firstView.put(0, (byte) 9));
    firstView.get();
    assertEquals(3, inputs.standardInputSnapshot().orElseThrow().remaining());
    assertEquals(1, opens.get());
    assertEquals(1, closes.get());
  }

  @Test
  void fileAliasesShareOneChargedSnapshotAndAggregateOverflowClosesPayload() throws Exception {
    Path source = Files.write(root.resolve("input.bin"), new byte[] {1, 2, 3});
    var authority =
        grant(
            5,
            List.of(
                new ExecutionGrant.ReadAuthority.File(source),
                new ExecutionGrant.ReadAuthority.Stdin()));
    var snapshots = new InputSnapshots(authority);
    var closes = new AtomicInteger();
    try (var access =
        new RequestPathAccess(
            root,
            dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory.rooted(
                    root.resolve("scratch"))
                ::createTempFile,
            authority,
            snapshots)) {
      Path first = access.materializeRead("input.bin", "payload", "input-", ".bin");
      assertEquals(first, access.materializeRead("./input.bin", "payload", "input-", ".bin"));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              snapshots.standardInput(() ->
                  new ByteArrayInputStream(new byte[] {4, 5, 6}) {
                    @Override
                    public void close() {
                      closes.incrementAndGet();
                    }
                  }));
      assertEquals(1, closes.get());
    }
    assertArrayEquals(new byte[] {1, 2, 3}, Files.readAllBytes(source));
  }

  @Test
  void overflowCannotRecoverAllowanceForAnotherReadOrPayloadOpener() throws Exception {
    var snapshots = new InputSnapshots(grant(3, List.of()));
    var bytesRead = new AtomicInteger();
    try (var first = countedInput(new byte[9], bytesRead)) {
      assertThrows(
          IllegalArgumentException.class,
          () -> snapshots.copy(first, OutputStream.nullOutputStream()));
    }
    assertEquals(4, bytesRead.get(), "Only the remaining allowance plus one lookahead is read");
    try (var next = countedInput(new byte[3], bytesRead)) {
      assertThrows(
          IllegalArgumentException.class,
          () -> snapshots.copy(next, OutputStream.nullOutputStream()));
    }
    assertEquals(4, bytesRead.get());
    var opens = new AtomicInteger();
    assertThrows(
        IllegalArgumentException.class,
        () ->
            snapshots.standardInput(() -> {
              opens.incrementAndGet();
              return new ByteArrayInputStream(new byte[3]);
            }));
    assertEquals(0, opens.get());
  }

  @Test
  void overflowBlocksNewFilesButPreservesPreviouslyCompletedAliases() throws Exception {
    Path cached = Files.write(root.resolve("cached.bin"), new byte[1]);
    Path overflow = Files.write(root.resolve("overflow.bin"), new byte[8]);
    Path later = Files.write(root.resolve("later.bin"), new byte[1]);
    var authority =
        grant(
            3,
            List.of(
                new ExecutionGrant.ReadAuthority.File(cached),
                new ExecutionGrant.ReadAuthority.File(overflow),
                new ExecutionGrant.ReadAuthority.File(later)));
    var creations = new AtomicInteger();
    try (var access =
        new RequestPathAccess(
            root,
            (prefix, suffix) -> {
              creations.incrementAndGet();
              return Files.createTempFile(root, prefix, suffix);
            },
            authority)) {
      Path snapshot = access.materializeRead("cached.bin", "payload", "snapshot-", ".bin");
      assertThrows(
          IllegalArgumentException.class,
          () -> access.materializeRead("overflow.bin", "payload", "snapshot-", ".bin"));
      assertThrows(
          IllegalArgumentException.class,
          () -> access.materializeRead("later.bin", "payload", "snapshot-", ".bin"));
      assertThrows(
          IllegalArgumentException.class,
          () -> access.materializeRead("overflow.bin", "payload", "snapshot-", ".bin"));
      assertEquals(2, creations.get(), "No file capability/temp acquisition after overflow");
      assertEquals(
          snapshot, access.materializeRead("./cached.bin", "payload", "snapshot-", ".bin"));
    }
  }

  @Test
  void partialReadFailureCannotStartAnotherAcquisitionButKeepsCachedStdin() throws Exception {
    var snapshots = new InputSnapshots(grant(5, List.of()));
    assertEquals(
        1, snapshots.standardInput(() -> new ByteArrayInputStream(new byte[1])).remaining());
    var reads = new AtomicInteger();
    try (InputStream broken =
        new InputStream() {
          @Override
          public int read() throws IOException {
            if (reads.incrementAndGet() > 2) throw new IOException("synthetic read failure");
            return 1;
          }

          @Override
          public int read(byte[] data, int offset, int length) throws IOException {
            data[offset] = (byte) read();
            return 1;
          }
        }) {
      assertThrows(
          IOException.class, () -> snapshots.copy(broken, OutputStream.nullOutputStream()));
    }
    var nextReads = new AtomicInteger();
    assertThrows(
        IOException.class,
        () ->
            snapshots.copy(countedInput(new byte[1], nextReads), OutputStream.nullOutputStream()));
    assertEquals(0, nextReads.get());
    assertEquals(
        1,
        snapshots
            .standardInput(() -> {
              throw new AssertionError("Cached snapshot reopened");
            })
            .remaining());
  }

  @Test
  void preflightCollectsSiblingOriginsWithoutOpeningStdinAfterFileOverflow() throws Exception {
    Path file = Files.writeString(root.resolve("overflow.txt"), "overflow");
    var opens = new AtomicInteger();
    var inputs =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            new ExecutionGrant(
                3,
                List.of(
                    new ExecutionGrant.ReadAuthority.File(file),
                    new ExecutionGrant.ReadAuthority.Stdin()),
                List.of("SET_CELL"),
                new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                AcceptancePolicy.minimum()),
            Optional.of(() -> {
              opens.incrementAndGet();
              return new ByteArrayInputStream(new byte[] {'o', 'k'});
            }),
            Optional.empty());
    var steps = new ArrayList<WorkbookStep>();
    for (var source :
        List.of(
            TextSourceSpec.utf8File("overflow.txt"),
            TextSourceSpec.standardInput(),
            TextSourceSpec.utf8File("overflow.txt"))) {
      steps.add(
          new MutationStep(
              new StepId("write-" + steps.size()),
              new CellSelector.ByAddress(new SheetName("Data"), new CellAddress("A1")),
              new CellMutationAction.SetCell(new CellValueSpec.Text(source))));
    }
    var plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            steps);
    try {
      var result = RequestPreflight.verify(plan, inputs);
      assertTrue(result.preparedRequest().isEmpty());
      assertEquals(0, opens.get());
      assertEquals(
          List.of(
              "steps[0].action.value.value.path",
              "steps[1].action.value.value.type",
              "steps[2].action.value.value.path"),
          result.problems().stream()
              .map(problem ->
                  assertInstanceOf(ProblemContext.ResolveInputs.class, problem.context()))
              .map(context -> context.json().orElseThrow().jsonPathValue().orElseThrow())
              .sorted()
              .toList());
    } finally {
      inputs.closeRunResources();
    }
  }

  private static InputStream countedInput(byte[] bytes, AtomicInteger bytesRead) {
    return new InputStream() {
      private final ByteArrayInputStream data = new ByteArrayInputStream(bytes);

      @Override
      public int read() {
        int next = data.read();
        if (next >= 0) bytesRead.incrementAndGet();
        return next;
      }

      @Override
      public int read(byte[] target, int offset, int length) {
        int count = data.read(target, offset, length);
        if (count > 0) bytesRead.addAndGet(count);
        return count;
      }

      @Override
      public void close() throws IOException {
        data.close();
      }
    };
  }

  @Test
  void suppliedBytesAreChargedAndFailedReadsReleaseOwnedStreams() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExecutionInputBindings(
                root,
                root.resolve("scratch"),
                new byte[] {1, 2, 3},
                grant(2, List.of(new ExecutionGrant.ReadAuthority.Stdin()))));
  }
}
