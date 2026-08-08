package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import java.io.IOException;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/**
 * Tests the specific filesystem facts reported when a requested response file cannot be written.
 */
class CliResponseTransportSupportTest {
  @Test
  void responseWriteMessagesClassifyEachFilesystemCauseWithoutLosingTheRequestedPath()
      throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-response-transport-directory-");
    Path existingFile = Files.createTempFile("gridgrind-response-transport-file-", ".json");
    Path target = directory.resolve("response.json");

    assertEquals(
        "Could not write response file " + target + ": permission denied",
        CliResponseTransportSupport.responseWriteMessage(
            new AccessDeniedException(target.toString()), target));
    assertEquals(
        "Could not write response file " + directory + ": Is a directory",
        CliResponseTransportSupport.responseWriteMessage(
            new FileAlreadyExistsException(directory.toString()), directory));
    assertEquals(
        "Could not write response file "
            + existingFile
            + ": already exists; GridGrind never replaces an existing response file implicitly",
        CliResponseTransportSupport.responseWriteMessage(
            new FileAlreadyExistsException(existingFile.toString()), existingFile));
    assertEquals(
        "Could not write response file " + target + ": cross-device link",
        CliResponseTransportSupport.responseWriteMessage(
            new FileSystemException(target.toString(), null, "cross-device link"), target));
    assertEquals(
        "Could not write response file " + target + ": conflict with existing-response.json",
        CliResponseTransportSupport.responseWriteMessage(
            new FileSystemException(target.toString(), "existing-response.json", null), target));
    assertEquals(
        "Could not write response file " + target,
        CliResponseTransportSupport.responseWriteMessage(
            new FileSystemException(target.toString(), null, null), target));
    assertEquals(
        "Could not write response file " + target,
        CliResponseTransportSupport.responseWriteMessage(
            new FileSystemException(target.toString(), "   ", "   "), target));
    assertEquals(
        "Could not write response file " + target,
        CliResponseTransportSupport.responseWriteMessage(new IOException(), target));
    assertEquals(
        "Could not write response file " + target,
        CliResponseTransportSupport.responseWriteMessage(new IOException("   "), target));
    assertEquals(
        "Could not write response file " + target + ": disk full",
        CliResponseTransportSupport.responseWriteMessage(new IOException("disk full"), target));
  }

  @Test
  void writeResponseProblemPreservesTheFilesystemMessageInTheProblemAndCause() throws IOException {
    Path target = Files.createTempDirectory("gridgrind-response-transport-problem-");

    var problem =
        CliResponseTransportSupport.writeResponseProblem(
            new FileAlreadyExistsException(target.toString()), target);

    assertEquals(GridGrindProblemCode.IO_ERROR, problem.code());
    assertEquals(problem.message(), problem.causes().getFirst().message());
    assertEquals(problem.context().stage(), problem.causes().getFirst().stage());
  }
}
