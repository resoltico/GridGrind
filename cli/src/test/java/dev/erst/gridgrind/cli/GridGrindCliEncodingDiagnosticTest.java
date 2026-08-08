package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import org.junit.jupiter.api.Test;

/** Verifies the request-byte encoding diagnostic exposed by the executable CLI surface. */
class GridGrindCliEncodingDiagnosticTest extends GridGrindCliTestSupport {
  @Test
  void returnsInvalidEncodingForMalformedUtf8() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(
                    new byte[] {'{', '"', 'x', '"', ':', (byte) 0xC3, (byte) 0x28}),
                stdout,
                stderr);

    CommandError failure = commandErrorOnStdout(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ENCODING, failure.primaryProblem().code());
    assertEquals("execute", failure.command());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonPath());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonLine());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonColumn());
  }
}
