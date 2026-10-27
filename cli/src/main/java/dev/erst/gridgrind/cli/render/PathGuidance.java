package dev.erst.gridgrind.cli.render;

/** Host file, stdin and scratch policy shared by CLI parsing and help. */
public final class PathGuidance {
  private PathGuidance() {}

  /** Describes the CLI-owned path or input policy. */
  public static String standardInputRequiresRequestMessage() {
    return "STANDARD_INPUT-authored values require --request so stdin is available for input"
        + " content instead of the request JSON";
  }

  /** Describes the CLI-owned path or input policy. */
  public static String stdinExecutionRootRequiredMessage() {
    return "Requests read from stdin require --execution-root so request-owned paths resolve"
        + " from one explicit directory";
  }

  /** Describes the CLI-owned path or input policy. */
  public static String requestOwnedPathResolutionSummary() {
    return "When the CLI reads a request via --request, relative request-owned paths resolve from"
        + " the request file directory. When the request JSON arrives on stdin, pass"
        + " --execution-root <path> and relative request-owned paths resolve from that"
        + " directory.";
  }

  /** Describes the CLI-owned path or input policy. */
  public static String cliFlagPathResolutionSummary() {
    return "--request and --response resolve from the current working directory, as do"
        + " --execution-root and --temp-root.";
  }

  /** Describes the CLI-owned path or input policy. */
  public static String cliScratchSpaceSummary() {
    return "GridGrind creates one private per-run scratch directory. Without --temp-root, that"
        + " scratch directory lives under the OS temporary-file root. With --temp-root <path>,"
        + " GridGrind creates the private scratch directory under that supplied parent path."
        + " Best-effort cleanup removes the managed scratch directory on normal command"
        + " completion.";
  }
}
