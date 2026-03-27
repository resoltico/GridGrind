package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.PrintWriter;
import java.io.StringWriter;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

/** Covers argument parsing and launcher exit semantics for the Jazzer harness runner. */
class JazzerHarnessRunnerTest {
  @Nested
  class ParseClassName {
    @Test
    void parseClassName_returnsSelectedClassName_whenArgumentsAreValid() {
      assertEquals(
          "dev.erst.gridgrind.jazzer.protocol.ProtocolRequestFuzzTest",
          JazzerHarnessRunner.parseClassName(
              new String[] {"--class", "dev.erst.gridgrind.jazzer.protocol.ProtocolRequestFuzzTest"}));
    }

    @Test
    void parseClassName_throwsWhenArgumentsAreMissing() {
      assertThrows(IllegalArgumentException.class, () -> JazzerHarnessRunner.parseClassName(new String[0]));
    }

    @Test
    void parseClassName_throwsWhenClassNameIsBlank() {
      assertThrows(
          IllegalArgumentException.class, () -> JazzerHarnessRunner.parseClassName(new String[] {"--class", " "}));
    }
  }

  @Nested
  class Run {
    @Test
    void run_returnsSuccess_whenTaggedJazzerTestsExist() {
      StringWriter errors = new StringWriter();

      int exitCode =
          JazzerHarnessRunner.run(
              SuccessfulTaggedHarness.class.getName(), new PrintWriter(errors, true));

      assertEquals(0, exitCode);
      assertTrue(errors.toString().isBlank());
    }

    @Test
    void run_returnsFailure_whenNoTaggedJazzerTestsExist() {
      StringWriter errors = new StringWriter();

      int exitCode =
          JazzerHarnessRunner.run(UntaggedHarness.class.getName(), new PrintWriter(errors, true));

      assertEquals(1, exitCode);
      assertTrue(errors.toString().contains("No Jazzer tests were discovered"));
    }
  }

  @Tag("jazzer")
  static class SuccessfulTaggedHarness {
    @Test
    void succeeds() {}
  }

  static class UntaggedHarness {
    @Test
    void succeeds() {}
  }
}
