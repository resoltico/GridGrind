package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Locks the diagnostic input-reference fallback for source-resolution failures. */
class InputResolutionFailureTest {
  @Test
  void preservesConcreteFilePathsAndKindsWhenTheExceptionDoesNotOwnThem() {
    IllegalArgumentException failure = new IllegalArgumentException("failure");

    assertEquals(
        InputReference.path("source-backed text", "note.txt"),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new TextSourceInput.Utf8File("note.txt"))
            .input());
    assertEquals(
        InputReference.kind("source-backed text"),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new TextSourceInput.StandardInput())
            .input());
    assertEquals(
        InputReference.path("source-backed binary", "payload.bin"),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new BinarySourceInput.File("payload.bin"))
            .input());
    assertEquals(
        InputReference.kind("source-backed binary"),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new BinarySourceInput.StandardInput())
            .input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new TextSourceInput.Inline("text"))
            .input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new BinarySourceInput.InlineBase64("dGVzdA=="))
            .input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(failure, Optional.empty(), new Object()).input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(failure, Optional.empty(), null).input());
  }
}
