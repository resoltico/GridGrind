package dev.erst.gridgrind.cli.command;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parsed CLI command model for one GridGrind process invocation. */
public sealed interface CliCommand
    permits CliCommand.Help,
        CliCommand.Version,
        CliCommand.License,
        CliCommand.Schema,
        CliCommand.TemplateRequest,
        CliCommand.TemplateGrant,
        CliCommand.RecipeList,
        CliCommand.RecipeShow,
        CliCommand.RecipeMaterialize,
        CliCommand.RecipeMatch,
        CliCommand.Doctor,
        CliCommand.Catalog,
        CliCommand.Run {
  /** One explicit help surface published by the CLI. */
  enum HelpTopic {
    OVERVIEW,
    PROTOCOL,
    GUIDANCE
  }

  /** Requests that help text be emitted as the primary command output. */
  public record Help(HelpTopic topic, Optional<Path> responsePath) implements CliCommand {
    public Help {
      Objects.requireNonNull(topic, "topic must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the version string be emitted as the primary command output. */
  public record Version(Optional<Path> responsePath) implements CliCommand {
    public Version {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the license text be emitted as the primary command output. */
  public record License(Optional<Path> responsePath) implements CliCommand {
    public License {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests the generated schema of one exchanged document. */
  public record Schema(
      dev.erst.gridgrind.json.schema.JsonSchemas.Document document, Optional<Path> responsePath)
      implements CliCommand {
    public Schema {
      Objects.requireNonNull(document, "document must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that a minimal valid request JSON document be emitted as the primary output. */
  public record TemplateRequest(Optional<Path> responsePath) implements CliCommand {
    public TemplateRequest {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests the minimal fail-closed authority document. */
  public record TemplateGrant(Optional<Path> responsePath) implements CliCommand {
    public TemplateGrant {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the unified machine-readable recipe catalog be emitted. */
  public record RecipeList(Optional<Path> responsePath) implements CliCommand {
    public RecipeList {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that one built-in runnable recipe request be emitted as the primary output. */
  public record RecipeShow(String lookupId, boolean details, Optional<Path> responsePath)
      implements CliCommand {
    public RecipeShow {
      Objects.requireNonNull(lookupId, "lookupId must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /**
   * Requests that one recipe request and its declared assets be published into one new workspace.
   */
  public record RecipeMaterialize(String lookupId, Path workspacePath) implements CliCommand {
    public RecipeMaterialize {
      Objects.requireNonNull(lookupId, "lookupId must not be null");
      Objects.requireNonNull(workspacePath, "workspacePath must not be null");
    }
  }

  /** Requests that one machine-readable recipe keyword match report be emitted. */
  public record RecipeMatch(String query, Optional<Path> responsePath) implements CliCommand {
    public RecipeMatch {
      Objects.requireNonNull(query, "query must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that one authored request be linted and summarized without execution. */
  public record Doctor(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      Optional<Path> grantPath,
      Optional<Path> secretsProviderPath,
      Optional<Path> responsePath)
      implements CliCommand {
    public Doctor {
      Objects.requireNonNull(requestPath, "requestPath must not be null");
      Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
      Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
      Objects.requireNonNull(grantPath, "grantPath must not be null");
      Objects.requireNonNull(secretsProviderPath, "secretsProviderPath must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the machine-readable protocol catalog be emitted as the primary output. */
  sealed interface Catalog extends CliCommand permits CatalogIndex, CatalogShow, CatalogSearch {
    /** Optional output path for the primary command payload. */
    public Optional<Path> responsePath();
  }

  /** Requests that the compact protocol-catalog index be emitted. */
  public record CatalogIndex(Optional<Path> responsePath) implements Catalog {
    public CatalogIndex {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /**
   * Requests that one uniquely matching catalog entry or type group be emitted.
   *
   * <p>Duplicate ids must be qualified as {@code <group>:<id>}.
   */
  public record CatalogShow(String lookupId, Optional<Path> responsePath) implements Catalog {
    public CatalogShow {
      Objects.requireNonNull(lookupId, "lookupId must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that a ranked protocol-catalog discovery report be emitted. */
  public record CatalogSearch(String searchQuery, Optional<Path> responsePath) implements Catalog {
    public CatalogSearch {
      Objects.requireNonNull(searchQuery, "searchQuery must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests protocol execution using stdin/stdout or explicit request/response file paths. */
  public record Run(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      Optional<Path> grantPath,
      Optional<Path> secretsProviderPath,
      Optional<Path> responsePath)
      implements CliCommand {
    public Run {
      Objects.requireNonNull(requestPath, "requestPath must not be null");
      Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
      Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
      Objects.requireNonNull(grantPath, "grantPath must not be null");
      Objects.requireNonNull(secretsProviderPath, "secretsProviderPath must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }
}
