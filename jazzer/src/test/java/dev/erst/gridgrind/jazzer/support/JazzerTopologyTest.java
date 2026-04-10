package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies the shared Jazzer topology contract consumed by runtime code and build logic. */
class JazzerTopologyTest {
  @Test
  void harnessValues_followCommittedTopologyOrder() {
    assertArrayEquals(
        new JazzerHarness[] {
          JazzerHarness.protocolRequest(),
          JazzerHarness.protocolWorkflow(),
          JazzerHarness.engineCommandSequence(),
          JazzerHarness.xlsxRoundTrip()
        },
        JazzerHarness.values());
  }

  @Test
  void runTargets_followCommittedTopologyOrder() {
    assertArrayEquals(
        new JazzerRunTarget[] {
          JazzerRunTarget.regression(),
          JazzerRunTarget.protocolRequest(),
          JazzerRunTarget.protocolWorkflow(),
          JazzerRunTarget.engineCommandSequence(),
          JazzerRunTarget.xlsxRoundTrip()
        },
        JazzerRunTarget.values());
  }

  @Test
  void harnessMetadata_matchesCommittedPaths() {
    assertEquals(
        "dev/erst/gridgrind/jazzer/protocol/ProtocolRequestFuzzTestInputs/readRequest",
        JazzerHarness.protocolRequest().inputResourceDirectory());
    assertEquals(
        "dev/erst/gridgrind/jazzer/engine/XlsxRoundTripFuzzTestInputs/roundTrip",
        JazzerHarness.xlsxRoundTrip().inputResourceDirectory());
  }

  @Test
  void runTargets_resolveStableTaskNamesAndHarnessAssignments() {
    assertEquals(
        JazzerRunTarget.protocolWorkflow(), JazzerRunTarget.fromTaskName("fuzzProtocolWorkflow"));
    assertEquals(
        List.of(
            JazzerHarness.protocolRequest(),
            JazzerHarness.protocolWorkflow(),
            JazzerHarness.engineCommandSequence(),
            JazzerHarness.xlsxRoundTrip()),
        JazzerRunTarget.regression().harnesses());
    assertEquals(
        JazzerHarness.protocolRequest(), JazzerRunTarget.protocolRequest().replayHarness());
  }

  @Test
  void invalidLookup_rejectsUnknownKeys() {
    assertThrows(IllegalArgumentException.class, () -> JazzerHarness.fromKey("missing"));
    assertThrows(IllegalArgumentException.class, () -> JazzerRunTarget.fromKey("missing"));
    assertThrows(IllegalArgumentException.class, () -> JazzerRunTarget.fromTaskName("missing"));
  }

  @Test
  void activeFuzzTargets_shareTheirKeyWithSingleHarness() {
    List<JazzerRunTarget> activeTargets =
        Arrays.stream(JazzerRunTarget.values()).filter(JazzerRunTarget::activeFuzzing).toList();

    for (JazzerRunTarget target : activeTargets) {
      assertEquals(1, target.harnesses().size());
      assertEquals(target.key(), target.harnesses().getFirst().key());
    }
  }
}
