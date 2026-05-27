package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.JazzerRunTarget;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Implements the local operator command surface that surrounds the Jazzer harnesses. */
public final class JazzerCli {
  private JazzerCli() {}

  /** Dispatches one local Jazzer operator command. */
  public static void main(String[] arguments) throws IOException {
    int exitCode = run(arguments, System.out, System.err);
    if (exitCode != 0) {
      System.exit(exitCode);
    }
  }

  /** Runs one local Jazzer operator command against injected text streams. */
  static int run(String[] arguments, PrintStream standardOutput, PrintStream errorOutput)
      throws IOException {
    Objects.requireNonNull(arguments, "arguments must not be null");
    Objects.requireNonNull(standardOutput, "standardOutput must not be null");
    Objects.requireNonNull(errorOutput, "errorOutput must not be null");
    try {
      return dispatch(Arrays.asList(arguments), standardOutput);
    } catch (IllegalArgumentException exception) {
      String subcommand = arguments.length == 0 ? "" : arguments[0];
      errorOutput.println(exception.getMessage());
      errorOutput.println();
      errorOutput.println(usageText(subcommand));
      return 2;
    }
  }

  private static int dispatch(List<String> args, PrintStream standardOutput) throws IOException {
    if (args.isEmpty()) {
      throw new IllegalArgumentException("A Jazzer subcommand is required.");
    }
    switch (args.getFirst()) {
      case "summarize-run" -> {
        summarizeRun(args.subList(1, args.size()), standardOutput);
        return 0;
      }
      case "status" -> {
        status(args.subList(1, args.size()), standardOutput);
        return 0;
      }
      case "report" -> {
        report(args.subList(1, args.size()), standardOutput);
        return 0;
      }
      case "list-findings" -> {
        listFindings(args.subList(1, args.size()), standardOutput);
        return 0;
      }
      case "list-corpus" -> {
        listCorpus(args.subList(1, args.size()), standardOutput);
        return 0;
      }
      case "replay" -> {
        return replay(args.subList(1, args.size()), standardOutput);
      }
      case "promote" -> {
        promote(args.subList(1, args.size()), standardOutput);
        return 0;
      }
      case "refresh-promoted-metadata" -> {
        refreshPromotedMetadata(args.subList(1, args.size()), standardOutput);
        return 0;
      }
      default ->
          throw new IllegalArgumentException("Unknown Jazzer subcommand: " + args.getFirst());
    }
  }

  private static void summarizeRun(List<String> args, PrintStream standardOutput)
      throws IOException {
    Path projectDirectory = requiredPath(args, "--project-dir");
    JazzerRunTarget target = JazzerRunTarget.fromKey(requiredValue(args, "--target"));
    LocalRunSummary summary =
        JazzerReportSupport.summarizeRun(
            projectDirectory,
            target,
            requiredValue(args, "--task"),
            requiredPath(args, "--log"),
            requiredPath(args, "--history"),
            requiredValue(args, "--started-at"),
            requiredValue(args, "--finished-at"),
            Integer.parseInt(requiredValue(args, "--exit-code")),
            new CorpusStats(
                Long.parseLong(requiredValue(args, "--corpus-before-files")),
                Long.parseLong(requiredValue(args, "--corpus-before-bytes"))));
    standardOutput.println(JazzerTextRenderer.renderSummary(summary));
  }

  private static void status(List<String> args, PrintStream standardOutput) throws IOException {
    Path projectDirectory = projectDirectory(args);
    Optional<String> targetKey = optionalValue(args, "--target");
    List<LocalRunSummary> summaries =
        targetKey.isEmpty()
            ? availableSummaries(projectDirectory)
            : singleSummary(projectDirectory, JazzerRunTarget.fromKey(targetKey.orElseThrow()));
    if (hasFlag(args, "--json")) {
      standardOutput.println(JazzerJson.toJson(new SummaryPayload(summaries)));
      return;
    }
    standardOutput.println(JazzerTextRenderer.renderStatus(summaries));
  }

  private static void report(List<String> args, PrintStream standardOutput) throws IOException {
    Path projectDirectory = projectDirectory(args);
    Optional<String> targetKey = optionalValue(args, "--target");
    List<LocalRunSummary> summaries =
        targetKey.isEmpty()
            ? availableSummaries(projectDirectory)
            : singleSummary(projectDirectory, JazzerRunTarget.fromKey(targetKey.orElseThrow()));
    if (hasFlag(args, "--json")) {
      standardOutput.println(JazzerJson.toJson(new SummaryPayload(summaries)));
      return;
    }
    if (summaries.isEmpty()) {
      standardOutput.println("No Jazzer summaries recorded yet.");
      return;
    }
    for (int index = 0; index < summaries.size(); index++) {
      if (index > 0) {
        standardOutput.println();
      }
      standardOutput.println(JazzerTextRenderer.renderSummary(summaries.get(index)));
    }
  }

  private static void listFindings(List<String> args, PrintStream standardOutput)
      throws IOException {
    Path projectDirectory = projectDirectory(args);
    Optional<String> targetKey = optionalValue(args, "--target");
    List<JazzerRunTarget> targets =
        targetKey.isEmpty()
            ? List.of(JazzerRunTarget.values())
            : List.of(JazzerRunTarget.fromKey(targetKey.orElseThrow()));
    if (hasFlag(args, "--json")) {
      ArrayList<TargetFindings> listings = new ArrayList<>();
      for (JazzerRunTarget target : targets) {
        listings.add(
            new TargetFindings(
                target.key(),
                JazzerReportSupport.findingArtifacts(target.workingDirectory(projectDirectory))));
      }
      standardOutput.println(JazzerJson.toJson(new FindingsPayload(List.copyOf(listings))));
      return;
    }

    for (int index = 0; index < targets.size(); index++) {
      JazzerRunTarget target = targets.get(index);
      if (index > 0) {
        standardOutput.println();
      }
      standardOutput.println(
          JazzerTextRenderer.renderFindingListing(
              target.key(),
              JazzerReportSupport.findingArtifacts(target.workingDirectory(projectDirectory))));
    }
  }

  private static void listCorpus(List<String> args, PrintStream standardOutput) throws IOException {
    Path projectDirectory = projectDirectory(args);
    Optional<String> targetKey = optionalValue(args, "--target");
    List<JazzerRunTarget> targets =
        targetKey.isEmpty()
            ? List.of(
                JazzerRunTarget.protocolRequest(),
                JazzerRunTarget.protocolWorkflow(),
                JazzerRunTarget.engineCommandSequence(),
                JazzerRunTarget.xlsxRoundTrip())
            : List.of(JazzerRunTarget.fromKey(targetKey.orElseThrow()));
    if (hasFlag(args, "--json")) {
      ArrayList<TargetCorpus> listings = new ArrayList<>();
      for (JazzerRunTarget target : targets) {
        List<Path> promotedInputs =
            target.replayable()
                ? JazzerReportSupport.promotedInputs(projectDirectory, target.replayHarness())
                : List.of();
        List<Path> orphans =
            target.replayable()
                ? JazzerReportSupport.orphanedInputs(projectDirectory, target.replayHarness())
                : List.of();
        listings.add(
            new TargetCorpus(
                target.key(),
                JazzerReportSupport.scanCorpus(target.workingDirectory(projectDirectory)),
                JazzerReportSupport.scanFiles(promotedInputs),
                JazzerReportSupport.newestCorpusEntries(
                    target.workingDirectory(projectDirectory), 10),
                promotedInputs,
                orphans));
      }
      standardOutput.println(JazzerJson.toJson(new CorpusPayload(List.copyOf(listings))));
      return;
    }

    for (int index = 0; index < targets.size(); index++) {
      JazzerRunTarget target = targets.get(index);
      List<Path> promotedInputs =
          target.replayable()
              ? JazzerReportSupport.promotedInputs(projectDirectory, target.replayHarness())
              : List.of();
      List<Path> orphans =
          target.replayable()
              ? JazzerReportSupport.orphanedInputs(projectDirectory, target.replayHarness())
              : List.of();
      if (index > 0) {
        standardOutput.println();
      }
      standardOutput.println(
          JazzerTextRenderer.renderCorpusListing(
              target.key(),
              JazzerReportSupport.scanCorpus(target.workingDirectory(projectDirectory)),
              JazzerReportSupport.scanFiles(promotedInputs),
              JazzerReportSupport.newestCorpusEntries(
                  target.workingDirectory(projectDirectory), 10),
              promotedInputs,
              orphans));
    }
  }

  private static int replay(List<String> args, PrintStream standardOutput) throws IOException {
    JazzerRunTarget target = JazzerRunTarget.fromKey(requiredValue(args, "--target"));
    if (!target.replayable()) {
      throw new IllegalArgumentException(
          "Replay requires a single-harness target, not " + target.key());
    }
    Path inputPath = requiredPath(args, "--input");
    byte[] input = readRequiredInputBytes(inputPath, "Replay input");
    ReplayOutcome outcome = JazzerReplaySupport.replay(target.replayHarness(), input);
    if (hasFlag(args, "--json")) {
      standardOutput.println(JazzerJson.toJson(outcome));
    } else {
      standardOutput.println(JazzerTextRenderer.renderReplay(inputPath, outcome));
    }
    return outcome instanceof ReplayOutcome.UnexpectedFailure ? 1 : 0;
  }

  private static void promote(List<String> args, PrintStream standardOutput) throws IOException {
    Path projectDirectory = projectDirectory(args);
    JazzerRunTarget target = JazzerRunTarget.fromKey(requiredValue(args, "--target"));
    if (!target.replayable()) {
      throw new IllegalArgumentException(
          "Promotion requires a single-harness target, not " + target.key());
    }

    Path inputPath = requiredPath(args, "--input");
    String name = requiredValue(args, "--name");
    byte[] input = readRequiredInputBytes(inputPath, "Promotion input");
    ReplayOutcome outcome = JazzerReplaySupport.replay(target.replayHarness(), input);

    Path promotedInputDirectory = target.replayHarness().inputDirectory(projectDirectory);
    Files.createDirectories(promotedInputDirectory);
    Path promotedInputPath = promotedInputDirectory.resolve(name + extensionFor(inputPath));
    Files.write(promotedInputPath, input);

    Path metadataDirectory = target.replayHarness().promotedMetadataDirectory(projectDirectory);
    Files.createDirectories(metadataDirectory);
    Path replayTextPath = metadataDirectory.resolve(name + ".txt");
    Path replayJsonPath = metadataDirectory.resolve(name + ".json");
    String storedSourcePath = PromotionMetadata.relativizePath(projectDirectory, inputPath);
    String storedPromotedInputPath =
        PromotionMetadata.relativizePath(projectDirectory, promotedInputPath);
    String storedReplayTextPath =
        PromotionMetadata.relativizePath(projectDirectory, replayTextPath);
    Files.writeString(
        replayTextPath,
        JazzerTextRenderer.renderReplay(Path.of(storedPromotedInputPath), outcome)
            + System.lineSeparator());
    JazzerJson.write(
        replayJsonPath,
        new PromotionMetadata(
            target.key(),
            storedSourcePath,
            storedPromotedInputPath,
            JazzerReplaySupport.outcomeKind(outcome),
            JazzerReplaySupport.expectationFor(outcome),
            Instant.now().toString(),
            storedReplayTextPath));

    standardOutput.println(
        "Promoted input written to " + promotedInputPath.toAbsolutePath().normalize());
    standardOutput.println(
        "Promotion metadata written to " + replayJsonPath.toAbsolutePath().normalize());
    if (outcome instanceof ReplayOutcome.UnexpectedFailure) {
      standardOutput.println("Promoted input currently reproduces an unexpected failure.");
    }
  }

  private static void refreshPromotedMetadata(List<String> args, PrintStream standardOutput)
      throws IOException {
    Path projectDirectory = projectDirectory(args);
    int refreshed =
        PromotionMetadataRefresher.refresh(projectDirectory, optionalValue(args, "--target"));
    standardOutput.println("Refreshed " + refreshed + " promoted metadata entries.");
  }

  private static List<LocalRunSummary> availableSummaries(Path projectDirectory)
      throws IOException {
    ArrayList<LocalRunSummary> summaries = new ArrayList<>();
    for (JazzerRunTarget target : JazzerRunTarget.values()) {
      if (JazzerReportSupport.hasLatestSummary(projectDirectory, target)) {
        summaries.add(JazzerReportSupport.readLatestSummary(projectDirectory, target));
      }
    }
    return List.copyOf(summaries);
  }

  private static List<LocalRunSummary> singleSummary(Path projectDirectory, JazzerRunTarget target)
      throws IOException {
    if (!JazzerReportSupport.hasLatestSummary(projectDirectory, target)) {
      return List.of();
    }
    return List.of(JazzerReportSupport.readLatestSummary(projectDirectory, target));
  }

  private static Path projectDirectory(List<String> args) {
    return optionalValue(args, "--project-dir")
        .map(Path::of)
        .map(Path::toAbsolutePath)
        .map(Path::normalize)
        .orElseGet(() -> Path.of("").toAbsolutePath().normalize());
  }

  private static Path requiredPath(List<String> args, String flag) {
    return Path.of(requiredValue(args, flag)).toAbsolutePath().normalize();
  }

  private static String requiredValue(List<String> args, String flag) {
    return optionalValue(args, flag)
        .orElseThrow(() -> new IllegalArgumentException("Missing required option " + flag));
  }

  private static Optional<String> optionalValue(List<String> args, String flag) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(flag, "flag must not be null");
    for (int index = 0; index < args.size() - 1; index++) {
      if (args.get(index).equals(flag)) {
        return Optional.of(args.get(index + 1));
      }
    }
    return Optional.empty();
  }

  private static boolean hasFlag(List<String> args, String flag) {
    return args.contains(flag);
  }

  private static byte[] readRequiredInputBytes(Path inputPath, String label) {
    Objects.requireNonNull(inputPath, "inputPath must not be null");
    Objects.requireNonNull(label, "label must not be null");
    if (!Files.exists(inputPath)) {
      throw new IllegalArgumentException(label + " does not exist: " + inputPath);
    }
    if (!Files.isRegularFile(inputPath)) {
      throw new IllegalArgumentException(label + " is not a regular file: " + inputPath);
    }
    try {
      return Files.readAllBytes(inputPath);
    } catch (IOException exception) {
      throw new IllegalArgumentException(
          label + " could not be read: " + inputPath + " (" + exception.getMessage() + ")",
          exception);
    }
  }

  private static String extensionFor(Path inputPath) {
    String fileName = inputPath.getFileName().toString();
    int separator = fileName.lastIndexOf('.');
    return separator >= 0 ? fileName.substring(separator) : ".bin";
  }

  private static String usageText(String subcommand) {
    String targets = String.join(", ", JazzerRunTarget.keys());
    return switch (subcommand) {
      case "status" ->
          "Usage: jazzer/bin/status [target] [--json] [gradle-options...]\nValid targets: "
              + targets;
      case "report" ->
          "Usage: jazzer/bin/report [target] [--json] [gradle-options...]\nValid targets: "
              + targets;
      case "list-findings" ->
          "Usage: jazzer/bin/list-findings [target] [--json] [gradle-options...]\nValid targets: "
              + targets;
      case "list-corpus" ->
          "Usage: jazzer/bin/list-corpus [target] [--json] [gradle-options...]\nValid targets: "
              + targets;
      case "replay" ->
          "Usage: jazzer/bin/replay <target> <input-path> [--json] [gradle-options...]\nValid targets: "
              + targets;
      case "promote" ->
          "Usage: jazzer/bin/promote <target> <input-path> <name> [gradle-options...]\nValid targets: "
              + targets;
      case "refresh-promoted-metadata" ->
          "Usage: jazzer/bin/refresh-promoted-metadata [target] [gradle-options...]\nValid targets: "
              + targets;
      case "summarize-run" ->
          "Usage: internal summarize-run requires --target, --task, --log, --history,"
              + " --started-at, --finished-at, --exit-code, --corpus-before-files,"
              + " and --corpus-before-bytes.";
      default ->
          "Usage: jazzer/bin/<command> [args]\nCommands: status, report, list-findings,"
              + " list-corpus, replay, promote, refresh-promoted-metadata.\nValid targets: "
              + targets;
    };
  }

  private record SummaryPayload(List<LocalRunSummary> summaries) {}

  private record FindingsPayload(List<TargetFindings> targets) {}

  private record TargetFindings(String target, List<FindingArtifact> findings) {}

  private record CorpusPayload(List<TargetCorpus> targets) {}

  private record TargetCorpus(
      String target,
      CorpusStats generatedLocalCorpus,
      CorpusStats committedCustomSeeds,
      List<Path> newestLocalCorpusEntries,
      List<Path> promotedInputs,
      List<Path> orphanedInputs) {}
}
