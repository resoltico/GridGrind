package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Host remediation is assembled at the CLI boundary instead of inferred in core. */
class ArgumentGuidanceTest {
  @Test
  void attachesCurrentHostSyntaxToTheRejectedArgument() {
    var request =
        CommandErrors.invalidArguments("run", Optional.of("--request"), "Missing request");
    assertTrue(request.primaryProblem().resolution().contains("--request -"));
    assertTrue(request.primaryProblem().resolution().contains("--execution-root"));
    var recipe = CommandErrors.invalidArguments("recipe show", Optional.of("id"), "Unknown recipe");
    assertTrue(recipe.primaryProblem().resolution().contains("recipe list"));
    var catalog = CommandErrors.invalidArguments("catalog show", Optional.of("id"), "Unknown id");
    assertTrue(catalog.primaryProblem().resolution().contains("catalog search"));
  }
}
