package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.io.RecipeWorkspacePublicationException;
import dev.erst.gridgrind.cli.io.RecipeWorkspacePublisher;
import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.core.discovery.RecipeCatalogs;
import dev.erst.gridgrind.core.discovery.RecipeMatcher;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Objects;
import java.util.Optional;

/** Built-in recipe-discovery CLI surfaces. */
final class RecipeCommand {
  private RecipeCommand() {}

  static int recipe(
      CliCommand.RecipeShow command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var recipe = RecipeRegistry.recipeFor(command.lookupId());
    if (recipe.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownRecipeMessage(command.lookupId());
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments("recipe show", Optional.of("id"), message),
          prettyJson);
    }
    var resolvedRecipe = recipe.get();
    if (command.details()) {
      return responseWriter.writeRenderedPayload(
          command.responsePath(),
          stdout,
          stderr,
          output ->
              GridGrindJsonOutput.writeCatalogLookupResult(
                  output,
                  RecipeCatalogs.catalog().protocolVersion(),
                  RecipeCatalogs.lookupFor(command.lookupId()).orElseThrow(),
                  prettyJson));
    }

    if (!resolvedRecipe.requiredWorkspacePaths().isEmpty()) {
      return responseWriter.writeCommandError(
          Optional.empty(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "recipe show",
              Optional.of("id"),
              "Asset-backed recipes require recipe materialize "
                  + resolvedRecipe.id()
                  + " --workspace <new-directory>."),
          prettyJson);
    }
    byte[] requestBytes = GridGrindJsonOutput.writeRequestBytes(resolvedRecipe.plan(), prettyJson);
    return responseWriter.writePayload(command.responsePath(), stdout, stderr, requestBytes, 0);
  }

  static int materializeRecipe(
      CliCommand.RecipeMaterialize command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    var recipe = RecipeRegistry.recipeFor(command.lookupId());
    if (recipe.isEmpty()) {
      return responseWriter.writeCommandError(
          Optional.empty(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "recipe materialize",
              Optional.of("id"),
              CliCatalogCommandSupport.unknownRecipeMessage(command.lookupId())),
          prettyJson);
    }
    var resolvedRecipe = recipe.orElseThrow();
    byte[] requestBytes = GridGrindJsonOutput.writeRequestBytes(resolvedRecipe.plan(), prettyJson);
    try {
      RecipeWorkspacePublisher.publish(resolvedRecipe, command.workspacePath(), requestBytes);
    } catch (RecipeWorkspacePublicationException exception) {
      return responseWriter.writeCommandError(
          Optional.empty(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "recipe materialize",
              Optional.of("--workspace"),
              Objects.requireNonNullElse(
                  exception.getMessage(), "Recipe workspace could not be published")),
          prettyJson);
    }
    responseWriter.writePayload(Optional.empty(), stdout, stderr, requestBytes, 0);
    return 0;
  }

  static int recipeCatalog(
      CliCommand.RecipeList command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return responseWriter.writePayload(
        command.responsePath(),
        stdout,
        stderr,
        GridGrindCliJson.writeBytes(RecipeCatalogs.catalog(), prettyJson),
        0);
  }

  static int recipeKeywordMatch(
      CliCommand.RecipeMatch command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    try {
      return responseWriter.writePayload(
          command.responsePath(),
          stdout,
          stderr,
          GridGrindCliJson.writeBytes(RecipeMatcher.reportFor(command.query()), prettyJson),
          0);
    } catch (IllegalArgumentException exception) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "recipe match",
              Optional.of("text"),
              Objects.requireNonNullElse(exception.getMessage(), "Invalid keyword query")),
          prettyJson);
    }
  }
}
