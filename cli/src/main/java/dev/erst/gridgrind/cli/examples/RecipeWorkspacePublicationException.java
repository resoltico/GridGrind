package dev.erst.gridgrind.cli.examples;

import java.io.IOException;

/** Reports one user-correctable inability to publish a recipe workspace atomically. */
public final class RecipeWorkspacePublicationException extends IOException {
  private static final long serialVersionUID = 1L;

  RecipeWorkspacePublicationException(String message) {
    super(message);
  }

  RecipeWorkspacePublicationException(String message, Throwable cause) {
    super(message, cause);
  }
}
