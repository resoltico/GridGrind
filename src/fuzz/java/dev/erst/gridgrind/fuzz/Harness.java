package dev.erst.gridgrind.fuzz;

import dev.erst.gridgrind.fuzz.engine.CommandSequenceHarness;
import dev.erst.gridgrind.fuzz.engine.XlsxRoundTripHarness;
import dev.erst.gridgrind.fuzz.protocol.ProtocolRequestHarness;
import dev.erst.gridgrind.fuzz.protocol.ProtocolWorkflowHarness;
import dev.erst.gridgrind.fuzz.support.GridGrindFuzzData;
import java.io.IOException;
import java.util.Arrays;

/** The same byte-level inputs and invariants drive deterministic replay and active fuzzing. */
public enum Harness {
  PROTOCOL_REQUEST("protocol-request"),
  PROTOCOL_WORKFLOW("protocol-workflow"),
  COMMAND_SEQUENCE("engine-command-sequence"),
  XLSX_ROUNDTRIP("xlsx-roundtrip");

  private final String key;

  Harness(String key) {
    this.key = key;
  }

  public String key() {
    return key;
  }

  public static Harness fromKey(String key) {
    return Arrays.stream(values())
        .filter(harness -> harness.key.equals(key))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown fuzz harness: " + key));
  }

  public void check(byte[] input) throws IOException {
    switch (this) {
      case PROTOCOL_REQUEST -> ProtocolRequestHarness.check(input);
      case PROTOCOL_WORKFLOW -> ProtocolWorkflowHarness.check(GridGrindFuzzData.replay(input));
      case COMMAND_SEQUENCE -> CommandSequenceHarness.check(GridGrindFuzzData.replay(input));
      case XLSX_ROUNDTRIP -> XlsxRoundTripHarness.check(GridGrindFuzzData.replay(input));
    }
  }
}
