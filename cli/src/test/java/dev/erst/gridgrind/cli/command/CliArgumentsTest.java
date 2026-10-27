package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Exercises the adopted command grammar and its input/output ownership boundaries. */
class CliArgumentsTest {
  @Test
  void parsesEachCommandWithoutFlagAliases() {
    assertInstanceOf(CliCommand.Help.class, CliArguments.parse(new String[] {"help"}));
    assertInstanceOf(CliCommand.Version.class, CliArguments.parse(new String[] {"version"}));
    assertInstanceOf(CliCommand.License.class, CliArguments.parse(new String[] {"license"}));
    assertInstanceOf(
        CliCommand.TemplateRequest.class, CliArguments.parse(new String[] {"template", "request"}));
    assertInstanceOf(
        CliCommand.TemplateGrant.class, CliArguments.parse(new String[] {"template", "grant"}));
    assertInstanceOf(
        CliCommand.RecipeList.class, CliArguments.parse(new String[] {"recipe", "list"}));
    assertInstanceOf(
        CliCommand.RecipeShow.class, CliArguments.parse(new String[] {"recipe", "show", "BUDGET"}));
    assertInstanceOf(
        CliCommand.RecipeMatch.class,
        CliArguments.parse(new String[] {"recipe", "match", "pivot chart"}));
    assertInstanceOf(CliCommand.CatalogIndex.class, CliArguments.parse(new String[] {"catalog"}));
    assertInstanceOf(
        CliCommand.CatalogSearch.class,
        CliArguments.parse(new String[] {"catalog", "search", "chart"}));
    assertInstanceOf(
        CliCommand.CatalogShow.class,
        CliArguments.parse(new String[] {"catalog", "show", "mutationActionTypes:SET_CELL"}));
    assertInstanceOf(
        CliCommand.Doctor.class,
        CliArguments.parse(new String[] {"doctor", "--request", "request.json"}));
  }

  @ParameterizedTest
  @ValueSource(strings = {"request", "grant", "response"})
  void parsesEachExchangedSchemaDocument(String document) {
    var command =
        assertInstanceOf(
            CliCommand.Schema.class, CliArguments.parse(new String[] {"schema", document}));
    assertEquals(document, command.document().name().toLowerCase(java.util.Locale.ROOT));
  }

  @Test
  void recipeDetailsAreAnExplicitViewOfTheSamePublishedRecipe() {
    var command =
        assertInstanceOf(
            CliCommand.RecipeShow.class,
            CliArguments.parse(new String[] {"recipe", "show", "CUSTOM_XML", "--details"}));
    assertTrue(command.details());
    assertThrows(
        CliArgumentsException.class,
        () -> CliArguments.parse(new String[] {"recipe", "list", "--details"}));
  }

  @Test
  void runRequiresExplicitRequestAndAuthority() {
    var command =
        assertInstanceOf(
            CliCommand.Run.class,
            CliArguments.parse(
                new String[] {
                  "run",
                  "--request",
                  "request.json",
                  "--grant",
                  "grant.json",
                  "--response",
                  "response.json"
                }));
    assertEquals(Path.of("request.json"), command.requestPath().orElseThrow());
    assertEquals(Path.of("grant.json"), command.grantPath().orElseThrow());
    assertThrows(
        CliArgumentsException.class,
        () -> CliArguments.parse(new String[] {"run", "--request", "request.json"}));
    assertThrows(
        CliArgumentsException.class,
        () -> CliArguments.parse(new String[] {"run", "--grant", "grant.json"}));
  }

  @ParameterizedTest
  @ValueSource(
      strings = {
        "--help",
        "-h",
        "--version",
        "--doctor-request",
        "--print-recipe",
        "--print-protocol-catalog"
      })
  void rejectsRetiredPrimaryFlagsRatherThanTranslatingThem(String token) {
    var failure =
        assertThrows(CliArgumentsException.class, () -> CliArguments.parse(new String[] {token}));
    assertTrue(failure.getMessage().startsWith("Unknown command: "));
  }

  @Test
  void rejectsDuplicateUnknownAndInapplicableOptions() {
    for (String[] args :
        List.of(
            new String[] {"schema", "request", "--response", "a.json", "--response", "b.json"},
            new String[] {"schema", "request", "--request", "input.json"},
            new String[] {"recipe", "list", "--lookup", "BUDGET"},
            new String[] {
              "run", "--request", "r.json", "--grant", "g.json", "--pretty", "--pretty"
            },
            new String[] {"help", "--response"},
            new String[] {"template", "response"},
            new String[] {"catalog", "show"},
            new String[] {"recipe", "show", ""})) {
      assertThrows(
          CliArgumentsException.class,
          () -> CliArguments.parse(args),
          java.util.Arrays.toString(args));
    }
  }

  @Test
  void stdinUsesAnExplicitRootAndCannotConsumeGrantOrResponseStdin() {
    assertThrows(
        CliArgumentsException.class,
        () -> CliArguments.parse(new String[] {"run", "--request", "-", "--grant", "grant.json"}));
    var command =
        assertInstanceOf(
            CliCommand.Run.class,
            CliArguments.parse(
                new String[] {
                  "run", "--request", "-", "--execution-root", "workspace", "--grant", "grant.json"
                }));
    assertEquals(Path.of("workspace"), command.executionRootPath().orElseThrow());
    assertThrows(
        CliArgumentsException.class,
        () ->
            CliArguments.parse(new String[] {"run", "--request", "request.json", "--grant", "-"}));
    assertThrows(
        CliArgumentsException.class,
        () -> CliArguments.parse(new String[] {"schema", "request", "--response", "-"}));
  }

  @Test
  void responseCannotReplaceTheRequestThroughEquivalentPathSpellings() {
    assertThrows(
        CliArgumentsException.class,
        () ->
            CliArguments.parse(
                new String[] {
                  "run",
                  "--request",
                  "request.json",
                  "--grant",
                  "grant.json",
                  "--response",
                  "request.json"
                }));
    assertThrows(
        CliArgumentsException.class,
        () ->
            CliArguments.parse(
                new String[] {
                  "run",
                  "--request",
                  "folder/../request.json",
                  "--grant",
                  "grant.json",
                  "--response",
                  "request.json"
                }));
  }

  @Test
  void materializationHasOneNewWorkspaceAndNoSeparateResponseTransport() {
    var command =
        assertInstanceOf(
            CliCommand.RecipeMaterialize.class,
            CliArguments.parse(
                new String[] {
                  "recipe", "materialize", "AUDIT_EXISTING_WORKBOOK", "--workspace", "workspace"
                }));
    assertEquals(Path.of("workspace"), command.workspacePath());
    assertThrows(
        CliArgumentsException.class,
        () -> CliArguments.parse(new String[] {"recipe", "materialize", "BUDGET"}));
    assertThrows(
        CliArgumentsException.class,
        () ->
            CliArguments.parse(
                new String[] {
                  "recipe",
                  "materialize",
                  "BUDGET",
                  "--workspace",
                  "workspace",
                  "--response",
                  "r.json"
                }));
  }

  @Test
  void formattingIsAHostOutputChoiceWithNoSecondCommand() {
    var invocation =
        CliArguments.parseInvocation(
            new String[] {"help", "protocol", "--format", "structured", "--pretty"});
    assertEquals(CliOutputFormat.STRUCTURED, invocation.outputFormat().orElseThrow());
    assertTrue(invocation.prettyJson());
    assertEquals(
        CliCommand.HelpTopic.PROTOCOL,
        assertInstanceOf(CliCommand.Help.class, invocation.command()).topic());
    assertThrows(
        CliArgumentsException.class,
        () -> CliArguments.parse(new String[] {"schema", "request", "--format", "text"}));
    assertThrows(
        CliArgumentsException.class, () -> CliArguments.parse(new String[] {"help", "version"}));
  }
}
