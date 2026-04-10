package dev.erst.gridgrind.jazzer.protocol;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import com.code_intelligence.jazzer.junit.FuzzTest;
import dev.erst.gridgrind.jazzer.support.HarnessTelemetry;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.SequenceIntrospection;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.json.GridGrindJson;
import dev.erst.gridgrind.protocol.json.InvalidJsonException;
import dev.erst.gridgrind.protocol.json.InvalidRequestException;
import dev.erst.gridgrind.protocol.json.InvalidRequestShapeException;
import java.io.IOException;

/** Fuzzes protocol request decoding and validation from raw JSON payloads. */
class ProtocolRequestFuzzTest {
  private static final HarnessTelemetry TELEMETRY =
      HarnessTelemetry.forHarness(JazzerHarness.protocolRequest());

  @FuzzTest
  void readRequest(FuzzedDataProvider data) throws IOException {
    byte[] bytes = data.consumeRemainingAsBytes();
    TELEMETRY.beginIteration(bytes.length);

    try {
      GridGrindRequest request = GridGrindJson.readRequest(bytes);
      if (request == null) {
        throw new IllegalStateException("readRequest returned null");
      }
      TELEMETRY.recordSourceKind(SequenceIntrospection.sourceKind(request));
      TELEMETRY.recordPersistenceKind(SequenceIntrospection.persistenceKind(request));
      TELEMETRY.recordSequenceKinds(SequenceIntrospection.operationKinds(request.operations()));
      TELEMETRY.recordReadKinds(SequenceIntrospection.readKinds(request.reads()));
      TELEMETRY.recordStyleKinds(SequenceIntrospection.styleKinds(request.operations()));
      TELEMETRY.recordSuccess();
    } catch (InvalidJsonException | InvalidRequestShapeException | InvalidRequestException expected) {
      // Expected invalid-payload classifications are the normal outcome for many fuzz inputs.
      TELEMETRY.recordExpectedInvalid(expected);
    } catch (IOException | RuntimeException unexpected) {
      TELEMETRY.recordUnexpectedFailure(unexpected);
      throw unexpected;
    }
  }
}
