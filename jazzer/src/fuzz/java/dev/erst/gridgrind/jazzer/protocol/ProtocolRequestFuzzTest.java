package dev.erst.gridgrind.jazzer.protocol;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import com.code_intelligence.jazzer.junit.FuzzTest;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.jazzer.support.HarnessTelemetry;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.SequenceIntrospection;
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
      WorkbookPlan request = GridGrindJson.readRequest(bytes);
      if (request == null) {
        throw new IllegalStateException("readRequest returned null");
      }
      TELEMETRY.recordSourceKind(SequenceIntrospection.sourceKind(request));
      TELEMETRY.recordPersistenceKind(SequenceIntrospection.persistenceKind(request));
      TELEMETRY.recordSequenceKinds(SequenceIntrospection.mutationKinds(request.mutationSteps()));
      TELEMETRY.recordAssertionKinds(
          SequenceIntrospection.assertionKinds(request.assertionSteps()));
      TELEMETRY.recordReadKinds(SequenceIntrospection.inspectionKinds(request.inspectionSteps()));
      TELEMETRY.recordStyleKinds(SequenceIntrospection.styleKinds(request.mutationSteps()));
      TELEMETRY.recordSuccess();
    } catch (InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException expected) {
      // Expected invalid-payload classifications are the normal outcome for many fuzz inputs.
      TELEMETRY.recordExpectedInvalid(expected);
    } catch (IOException | RuntimeException unexpected) {
      TELEMETRY.recordUnexpectedFailure(unexpected);
      throw unexpected;
    }
  }
}
