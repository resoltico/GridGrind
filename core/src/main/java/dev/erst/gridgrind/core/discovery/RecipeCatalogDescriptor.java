package dev.erst.gridgrind.core.discovery;

import java.util.List;

/** Shared stable fields published by both compact recipe rows and lookup detail payloads. */
public interface RecipeCatalogDescriptor {
  /** Stable upper-case published recipe identifier. */
  String id();

  /** Portable request file name emitted when the recipe is printed. */
  String requestFileName();

  /** Short summary of what the recipe demonstrates or enables. */
  String summary();

  /** Whether the recipe is self-contained or requires copied example assets. */
  RecipeAdvisory advisory();

  /** Workspace-relative asset paths required before executing an asset-backed recipe. */
  List<String> requiredWorkspacePaths();
}
