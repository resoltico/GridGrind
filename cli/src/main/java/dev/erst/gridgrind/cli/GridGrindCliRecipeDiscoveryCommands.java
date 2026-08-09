package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.GridGrindRecipeCatalog;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Objects;
import java.util.Optional;

/** Built-in recipe-discovery CLI surfaces. */
final class GridGrindCliRecipeDiscoveryCommands {
  private GridGrindCliRecipeDiscoveryCommands() {}

  static int recipe(
      CliCommand.PrintRecipe command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var recipe = GridGrindCliRecipeRegistry.recipeFor(command.lookupId());
    if (recipe.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownRecipeMessage(command.lookupId());
      return CliCatalogPayloadSupport.writeCommandError(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments("print-recipe", Optional.of("--lookup"), message),
          prettyJson);
    }
    byte[] requestBytes = GridGrindJsonOutput.writeRequestBytes(recipe.get().plan(), prettyJson);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "print-recipe",
        command.responsePath(),
        stdout,
        stderr,
        requestBytes,
        prettyJson);
  }

  static int recipeCatalog(
      CliCommand.PrintRecipeCatalog command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    if (command.lookupId().isEmpty()) {
      return CliCatalogPayloadSupport.writePayload(
          responseWriter,
          "print-recipe-catalog",
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeBytes(GridGrindRecipeCatalog.catalog(), prettyJson),
          prettyJson);
    }
    String recipeFilter = command.lookupId().orElseThrow();
    var entry = GridGrindRecipeCatalog.lookupFor(recipeFilter);
    if (entry.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownRecipeMessage(recipeFilter);
      return CliCatalogPayloadSupport.writeCommandError(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments("print-recipe-catalog", Optional.of("--lookup"), message),
          prettyJson);
    }
    return CliCatalogPayloadSupport.writeRenderedPayload(
        responseWriter,
        "print-recipe-catalog",
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJsonOutput.writeCatalogLookupResult(
                output,
                GridGrindRecipeCatalog.catalog().protocolVersion(),
                entry.get(),
                prettyJson),
        prettyJson);
  }

  static int recipeKeywordMatch(
      CliCommand.PrintRecipeKeywordMatch command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    try {
      return CliCatalogPayloadSupport.writePayload(
          responseWriter,
          "print-recipe-keyword-match",
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeBytes(
              GridGrindRecipeKeywordMatcher.reportFor(command.query()), prettyJson),
          prettyJson);
    } catch (IllegalArgumentException exception) {
      return CliCatalogPayloadSupport.writeCommandError(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "print-recipe-keyword-match",
              Optional.of("--query"),
              Objects.requireNonNullElse(exception.getMessage(), "Invalid keyword query")),
          prettyJson);
    }
  }
}
