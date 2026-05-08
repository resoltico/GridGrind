package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parsed CLI command model for one GridGrind process invocation. */
sealed interface CliCommand {
  /** Requests that help text be emitted as the primary command output. */
  record Help(Optional<Path> responsePath) implements CliCommand {
    public Help {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the version string be emitted as the primary command output. */
  record Version(Optional<Path> responsePath) implements CliCommand {
    public Version {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the license text be emitted as the primary command output. */
  record License(Optional<Path> responsePath) implements CliCommand {
    public License {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that a minimal valid request JSON document be emitted as the primary output. */
  record PrintRequestTemplate(Optional<Path> responsePath) implements CliCommand {
    public PrintRequestTemplate {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that one built-in generated example request be emitted as the primary output. */
  record PrintExample(String exampleId, Optional<Path> responsePath) implements CliCommand {
    public PrintExample {
      Objects.requireNonNull(exampleId, "exampleId must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /**
   * Requests that the machine-readable built-in example catalog be emitted as the primary output.
   */
  record PrintExampleCatalog(Optional<Path> responsePath) implements CliCommand {
    public PrintExampleCatalog {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the machine-readable task catalog be emitted as the primary output. */
  record PrintTaskCatalog(Optional<String> taskFilter, Optional<Path> responsePath)
      implements CliCommand {
    public PrintTaskCatalog {
      Objects.requireNonNull(taskFilter, "taskFilter must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that one machine-readable starter task plan be emitted as the primary output. */
  record PrintTaskPlan(String taskId, Optional<Path> responsePath) implements CliCommand {
    public PrintTaskPlan {
      Objects.requireNonNull(taskId, "taskId must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /**
   * Requests that one machine-readable task keyword match report be emitted as the primary output.
   */
  record PrintTaskKeywordMatch(String query, Optional<Path> responsePath) implements CliCommand {
    public PrintTaskKeywordMatch {
      Objects.requireNonNull(query, "query must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that one authored request be linted and summarized without execution. */
  record DoctorRequest(Optional<Path> requestPath, Optional<Path> responsePath)
      implements CliCommand {
    public DoctorRequest {
      Objects.requireNonNull(requestPath, "requestPath must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that the machine-readable protocol catalog be emitted as the primary output. */
  sealed interface PrintProtocolCatalog extends CliCommand
      permits PrintProtocolCatalogAll, PrintProtocolCatalogLookup, PrintProtocolCatalogSearch {
    /** Optional output path for the primary command payload. */
    Optional<Path> responsePath();
  }

  /** Requests that the full protocol catalog be emitted. */
  record PrintProtocolCatalogAll(Optional<Path> responsePath) implements PrintProtocolCatalog {
    public PrintProtocolCatalogAll {
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /**
   * Requests that one uniquely matching catalog entry or type group be emitted.
   *
   * <p>Duplicate ids must be qualified as {@code <group>:<id>}.
   */
  record PrintProtocolCatalogLookup(String operationFilter, Optional<Path> responsePath)
      implements PrintProtocolCatalog {
    public PrintProtocolCatalogLookup {
      Objects.requireNonNull(operationFilter, "operationFilter must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests that a ranked protocol-catalog discovery report be emitted. */
  record PrintProtocolCatalogSearch(String searchQuery, Optional<Path> responsePath)
      implements PrintProtocolCatalog {
    public PrintProtocolCatalogSearch {
      Objects.requireNonNull(searchQuery, "searchQuery must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }

  /** Requests protocol execution using stdin/stdout or explicit request/response file paths. */
  record Execute(Optional<Path> requestPath, Optional<Path> responsePath) implements CliCommand {
    public Execute {
      Objects.requireNonNull(requestPath, "requestPath must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
    }
  }
}
