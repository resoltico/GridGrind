package dev.erst.gridgrind.fuzz.protocol;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.InvalidEncodingException;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.io.IOException;

/** Fuzzes protocol request decoding and validation from raw JSON payloads. */
public final class ProtocolRequestHarness {
  private ProtocolRequestHarness() {}

  public static void check(byte[] bytes) throws IOException {

    try {
      WorkbookPlan request = GridGrindJson.readRequest(bytes);
      if (request == null) {
        throw new IllegalStateException("readRequest returned null");
      }
    } catch (InvalidEncodingException
        | InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException expected) {
      // Expected invalid-payload classifications are the normal outcome for many fuzz inputs.
    }
  }
}
