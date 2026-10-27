package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
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
                failure, Optional.empty(), new TextSourceSpec.Utf8File("note.txt"))
            .input());
    assertEquals(
        InputReference.kind("source-backed text"),
        InputResolutionFailure.forSource(failure, Optional.empty(), new TextSourceSpec.Stdin())
            .input());
    assertEquals(
        InputReference.path("source-backed binary", "payload.bin"),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new BinarySourceSpec.File("payload.bin"))
            .input());
    assertEquals(
        InputReference.kind("source-backed binary"),
        InputResolutionFailure.forSource(failure, Optional.empty(), new BinarySourceSpec.Stdin())
            .input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new TextSourceSpec.Inline("text"))
            .input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(
                failure, Optional.empty(), new BinarySourceSpec.InlineBase64("dGVzdA=="))
            .input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(failure, Optional.empty(), new Object()).input());
    assertEquals(
        InputReference.unknown(),
        InputResolutionFailure.forSource(failure, Optional.empty(), null).input());
  }
}
