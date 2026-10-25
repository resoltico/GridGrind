package dev.erst.gridgrind.jazzer.protocol;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import com.code_intelligence.jazzer.junit.FuzzTest;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.InvalidEncodingException;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.jazzer.support.HarnessTelemetry;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.SequenceIntrospection;
import dev.erst.gridgrind.json.codec.GridGrindJson;
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
      TELEMETRY.recordSequenceKinds(
          SequenceIntrospection.mutationKinds(request.stepPartition().mutations()));
      TELEMETRY.recordAssertionKinds(
          SequenceIntrospection.assertionKinds(request.stepPartition().assertions()));
      TELEMETRY.recordReadKinds(
          SequenceIntrospection.inspectionKinds(request.stepPartition().inspections()));
      TELEMETRY.recordStyleKinds(
          SequenceIntrospection.styleKinds(request.stepPartition().mutations()));
      TELEMETRY.recordSuccess();
    } catch (InvalidEncodingException
        | InvalidJsonException
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
