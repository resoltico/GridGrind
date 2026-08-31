package dev.erst.gridgrind.cli.examples;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;

/** Copies packaged recipe assets into one private staging directory owned by the publisher. */
final class RecipeAssetMaterializer {
  private static final String RESOURCE_ROOT = "gridgrind/recipe-assets/";

  private RecipeAssetMaterializer() {}

  /** Copies every declared recipe asset into one empty publisher-owned staging directory. */
  static void copyToStaging(GridGrindCliRecipe recipe, Path workspace) throws IOException {
    Objects.requireNonNull(recipe, "recipe must not be null");
    Objects.requireNonNull(workspace, "workspace must not be null");
    Path root = workspace.toAbsolutePath().normalize();
    for (String relativePath : recipe.requiredWorkspacePaths()) {
      Path destination = root.resolve(relativePath).normalize();
      if (!destination.startsWith(root)) {
        throw new IOException("Recipe asset path escapes workspace: " + relativePath);
      }
      Path parent =
          Objects.requireNonNull(destination.getParent(), "asset parent must not be null");
      java.nio.file.Files.createDirectories(parent);
      try (InputStream resource = resource(relativePath)) {
        java.nio.file.Files.copy(resource, destination);
      }
    }
  }

  private static InputStream resource(String relativePath) throws IOException {
    InputStream stream =
        Thread.currentThread()
            .getContextClassLoader()
            .getResourceAsStream(RESOURCE_ROOT + relativePath);
    if (stream == null) {
      throw new IOException("Packaged recipe asset is missing: " + relativePath);
    }
    return stream;
  }
}
