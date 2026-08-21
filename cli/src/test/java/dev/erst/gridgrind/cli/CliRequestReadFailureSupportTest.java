package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests the request-input identity used by pre-execution diagnostics. */
class CliRequestReadFailureSupportTest {
  @Test
  void requestInputTreatsImplicitAndExplicitStandardInputConsistently() {
    RequestInput.StandardInput implicit =
        assertInstanceOf(
            RequestInput.StandardInput.class,
            CliRequestReadFailureSupport.requestInput(Optional.empty()));
    RequestInput.StandardInput explicit =
        assertInstanceOf(
            RequestInput.StandardInput.class,
            CliRequestReadFailureSupport.requestInput(Optional.of(Path.of("-"))));
    RequestInput.RequestFile file =
        assertInstanceOf(
            RequestInput.RequestFile.class,
            CliRequestReadFailureSupport.requestInput(Optional.of(Path.of("request.json"))));

    assertEquals(implicit, explicit);
    assertEquals(Path.of("request.json").toAbsolutePath().toString(), file.requestPath());
  }
}
