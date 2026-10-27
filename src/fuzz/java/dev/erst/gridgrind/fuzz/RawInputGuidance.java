package dev.erst.gridgrind.fuzz;

import edu.berkeley.cs.jqf.fuzz.ei.ZestGuidance;
import edu.berkeley.cs.jqf.fuzz.guidance.Result;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.time.Duration;
import java.util.Date;
import java.util.Random;

/** Zest guidance over finite raw byte tapes, without a generator envelope or seed conversion. */
final class RawInputGuidance extends ZestGuidance {
  private Throwable firstFailure;
  private final int inputBound;
  private final int seedCount;

  RawInputGuidance(
      String name,
      Duration duration,
      long trials,
      File output,
      File findings,
      File[] seeds,
      int inputBound,
      Random random)
      throws IOException {
    super(name, duration, trials, output, random);
    if (inputBound <= 0) throw new IllegalArgumentException("Input bound must be positive");
    this.inputBound = inputBound;
    this.seedCount = seeds.length;
    savedFailuresDirectory = findings;
    for (File seed : seeds) {
      byte[] bytes = Files.readAllBytes(seed.toPath());
      if (bytes.length > inputBound)
        throw new IllegalArgumentException("Input bound excludes seed: " + seed);
      seedInputs.add(new BytesInput(bytes));
    }
  }

  @Override
  public boolean hasInput() {
    // Admit all seed trials like native corpus initialization, subject to the outer process bound.
    if (!seedInputs.isEmpty() && uniqueFailures.isEmpty() && numTrials < maxTrials) return true;
    return super.hasInput();
  }

  @Override
  public void beforeRun() {
    runStart = new Date();
    branchCount = 0;
  }

  @Override
  public void handleResult(Result result, Throwable error) {
    if ((result == Result.FAILURE || result == Result.TIMEOUT) && firstFailure == null)
      firstFailure = error;
    super.handleResult(result, error);
  }

  Throwable firstFailure() {
    return firstFailure;
  }

  long trials() {
    return numTrials;
  }

  long seedTrialsExecuted() {
    return Math.min(numTrials, seedCount);
  }

  int seedCount() {
    return seedCount;
  }

  int coverage() {
    return getTotalCoverage().getNonZeroCount();
  }

  @Override
  protected Input<?> createFreshInput() {
    byte[] bytes = new byte[1 + random.nextInt(Math.min(64, inputBound))];
    random.nextBytes(bytes);
    return new BytesInput(bytes);
  }

  /** Finite identity tape with native overwrite mutation plus bounded length changes. */
  private final class BytesInput extends LinearInput {
    BytesInput(byte[] bytes) {
      for (byte value : bytes) values.add(Byte.toUnsignedInt(value));
    }

    BytesInput(LinearInput input) {
      super(input);
    }

    @Override
    public int getOrGenerateFresh(Integer index, Random source) {
      if (index >= values.size()) return -1;
      if (index != requested)
        throw new IllegalStateException("Raw input bytes requested out of order");
      requested++;
      return values.get(index);
    }

    @Override
    public void gc() {
      if (!values.isEmpty()) super.gc();
    }

    @Override
    public Input<Integer> fuzz(Random source) {
      BytesInput child =
          values.isEmpty()
              ? new BytesInput(new byte[0])
              : new BytesInput((LinearInput) super.fuzz(source));
      int operation = source.nextInt(3);
      if ((operation == 0 || child.values.isEmpty()) && child.values.size() < inputBound) {
        child.values.add(source.nextInt(child.values.size() + 1), source.nextInt(256));
      } else if (operation == 1 && !child.values.isEmpty()) {
        child.values.remove(source.nextInt(child.values.size()));
      }
      return child;
    }
  }
}
