package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

class InstrumentationOutputTest {
  @Test
  void detectsSplitWarningAndPreservesOutput() throws Exception {
    var destination = new ByteArrayOutputStream();
    var output = new InstrumentationOutput(destination);
    byte[] prefix = "[INFO] Instrumenting: product... Done!\n[WARNING] Could not ".getBytes(StandardCharsets.UTF_8);
    output.write(prefix);
    assertFalse(output.getFailed());
    byte[] suffix = "instrument product\n".getBytes(StandardCharsets.UTF_8);
    for (byte value : suffix) output.write(value);
    assertTrue(output.getFailed());
    assertArrayEquals((new String(prefix, StandardCharsets.UTF_8) + new String(suffix, StandardCharsets.UTF_8)).getBytes(StandardCharsets.UTF_8), destination.toByteArray());
  }
}
