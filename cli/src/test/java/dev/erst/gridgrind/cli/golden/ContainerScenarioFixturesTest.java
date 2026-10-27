package dev.erst.gridgrind.cli.golden;

import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises smoke inputs locally before an opt-in image build spends Docker time on them. */
class ContainerScenarioFixturesTest {
  @Test
  void smokeRequestsExerciseTheActualHostAndBackends(@TempDir Path workspace) throws Exception {
    ContainerScenarios.verify(GoldenLauncher.IN_PROCESS, workspace);
  }
}
