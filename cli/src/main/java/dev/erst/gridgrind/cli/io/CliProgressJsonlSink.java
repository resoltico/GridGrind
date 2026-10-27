package dev.erst.gridgrind.cli.io;

import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.core.result.ExecutionProgressEvent;
import dev.erst.gridgrind.engine.api.GridGrindProgressSink;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Objects;

/** Renders live execution progress as one compact JSON object per stderr line. */
public final class CliProgressJsonlSink implements GridGrindProgressSink {
  private final OutputStream stderr;

  /** Routes progress events to the caller-owned diagnostic stream. */
  public CliProgressJsonlSink(OutputStream stderr) {
    this.stderr = Objects.requireNonNull(stderr, "stderr must not be null");
  }

  @Override
  public void emit(ExecutionProgressEvent event) {
    try {
      CliPayloadOutput.write(stderr, GridGrindCliJson.writeBytes(event));
    } catch (IOException exception) {
      throw new CliPrimaryOutputException(exception);
    }
  }
}
