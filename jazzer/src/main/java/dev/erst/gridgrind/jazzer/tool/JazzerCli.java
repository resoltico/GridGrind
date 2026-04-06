package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.JazzerRunTarget;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Implements the local operator command surface that surrounds the Jazzer harnesses. */
public final class JazzerCli {
  private JazzerCli() {}

  /** Dispatches one local Jazzer operator command. */
  public static void main(String[] arguments) throws Exception {
    if (arguments.length == 0) {
      throw new IllegalArgumentException("A Jazzer subcommand is required");
    }

    List<String> args = Arrays.asList(arguments);
    switch (args.getFirst()) {
      case "summarize-run" -> summarizeRun(args.subList(1, args.size()));
      case "status" -> status(args.subList(1, args.size()));
      case "report" -> report(args.subList(1, args.size()));
      case "list-findings" -> listFindings(args.subList(1, args.size()));
      case "list-corpus" -> listCorpus(args.subList(1, args.size()));
      case "replay" -> replay(args.subList(1, args.size()));
      case "promote" -> promote(args.subList(1, args.size()));
      case "refresh-promoted-metadata" -> refreshPromotedMetadata(args.subList(1, args.size()));
      default -> throw new IllegalArgumentException("Unknown Jazzer subcommand: " + args.getFirst());
    }
  }

  private static void summarizeRun(List<String> args) throws IOException {
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
    System.out.println(JazzerTextRenderer.renderSummary(summary));
  }

  private static void status(List<String> args) throws IOException {
    Path projectDirectory = projectDirectory(args);
    String targetKey = optionalValue(args, "--target");
    List<LocalRunSummary> summaries =
        targetKey == null
            ? availableSummaries(projectDirectory)
            : singleSummary(projectDirectory, JazzerRunTarget.fromKey(targetKey));
    if (summaries.isEmpty()) {
      System.out.println("No Jazzer summaries recorded yet.");
      return;
    }
    System.out.println(JazzerTextRenderer.renderStatus(summaries));
  }

  private static void report(List<String> args) throws IOException {
    Path projectDirectory = projectDirectory(args);
    String targetKey = optionalValue(args, "--target");
    List<LocalRunSummary> summaries =
        targetKey == null
            ? availableSummaries(projectDirectory)
            : singleSummary(projectDirectory, JazzerRunTarget.fromKey(targetKey));
    if (summaries.isEmpty()) {
      System.out.println("No Jazzer summaries recorded yet.");
      return;
    }
    for (int index = 0; index < summaries.size(); index++) {
      if (index > 0) {
        System.out.println();
      }
      System.out.println(JazzerTextRenderer.renderSummary(summaries.get(index)));
    }
  }

  private static void listFindings(List<String> args) throws IOException {
    Path projectDirectory = projectDirectory(args);
    String targetKey = optionalValue(args, "--target");
    List<JazzerRunTarget> targets =
        targetKey == null ? List.of(JazzerRunTarget.values()) : List.of(JazzerRunTarget.fromKey(targetKey));

    for (int index = 0; index < targets.size(); index++) {
      JazzerRunTarget target = targets.get(index);
      if (index > 0) {
        System.out.println();
      }
      System.out.println(
          JazzerTextRenderer.renderFindingListing(
              target.key(),
              JazzerReportSupport.findingArtifacts(target.workingDirectory(projectDirectory))));
    }
  }

  private static void listCorpus(List<String> args) throws IOException {
    Path projectDirectory = projectDirectory(args);
    String targetKey = optionalValue(args, "--target");
    List<JazzerRunTarget> targets =
        targetKey == null
            ? List.of(
                JazzerRunTarget.PROTOCOL_REQUEST,
                JazzerRunTarget.PROTOCOL_WORKFLOW,
                JazzerRunTarget.ENGINE_COMMAND_SEQUENCE,
                JazzerRunTarget.XLSX_ROUND_TRIP)
            : List.of(JazzerRunTarget.fromKey(targetKey));

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
        System.out.println();
      }
      System.out.println(
          JazzerTextRenderer.renderCorpusListing(
              target.key(),
              JazzerReportSupport.scanCorpus(target.workingDirectory(projectDirectory)),
              JazzerReportSupport.scanFiles(promotedInputs),
              JazzerReportSupport.newestCorpusEntries(target.workingDirectory(projectDirectory), 10),
              promotedInputs,
              orphans));
    }
  }

  private static void replay(List<String> args) throws IOException {
    Path projectDirectory = projectDirectory(args);
    JazzerRunTarget target = JazzerRunTarget.fromKey(requiredValue(args, "--target"));
    if (!target.replayable()) {
      throw new IllegalArgumentException("Replay requires a single-harness target, not " + target.key());
    }
    Path inputPath = requiredPath(args, "--input");
    ReplayOutcome outcome =
        JazzerReplaySupport.replay(target.replayHarness(), Files.readAllBytes(inputPath));
    if (hasFlag(args, "--json")) {
      System.out.println(JazzerJson.toJson(outcome));
    } else {
      System.out.println(JazzerTextRenderer.renderReplay(inputPath, outcome));
    }
    if (outcome instanceof ReplayOutcome.UnexpectedFailure) {
      System.exit(1);
    }
  }

  private static void promote(List<String> args) throws IOException {
    Path projectDirectory = projectDirectory(args);
    JazzerRunTarget target = JazzerRunTarget.fromKey(requiredValue(args, "--target"));
    if (!target.replayable()) {
      throw new IllegalArgumentException("Promotion requires a single-harness target, not " + target.key());
    }

    Path inputPath = requiredPath(args, "--input");
    String name = requiredValue(args, "--name");
    byte[] input = Files.readAllBytes(inputPath);
    ReplayOutcome outcome = JazzerReplaySupport.replay(target.replayHarness(), input);

    Path promotedInputDirectory = target.replayHarness().inputDirectory(projectDirectory);
    Files.createDirectories(promotedInputDirectory);
    Path promotedInputPath = promotedInputDirectory.resolve(name + extensionFor(inputPath));
    Files.write(promotedInputPath, input);

    Path metadataDirectory =
        projectDirectory
            .resolve("src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata")
            .resolve(target.key());
    Files.createDirectories(metadataDirectory);
    Path replayTextPath = metadataDirectory.resolve(name + ".txt");
    Path replayJsonPath = metadataDirectory.resolve(name + ".json");
    Files.writeString(replayTextPath, JazzerTextRenderer.renderReplay(inputPath, outcome) + System.lineSeparator());
    JazzerJson.write(
        replayJsonPath,
        new PromotionMetadata(
            target.key(),
            inputPath.toAbsolutePath().normalize().toString(),
            promotedInputPath.toAbsolutePath().normalize().toString(),
            JazzerReplaySupport.outcomeKind(outcome),
            JazzerReplaySupport.expectationFor(outcome),
            Instant.now().toString(),
            replayTextPath.toAbsolutePath().normalize().toString()));

    System.out.println("Promoted input written to " + promotedInputPath.toAbsolutePath().normalize());
    System.out.println("Promotion metadata written to " + replayJsonPath.toAbsolutePath().normalize());
    if (outcome instanceof ReplayOutcome.UnexpectedFailure) {
      System.out.println("Promoted input currently reproduces an unexpected failure.");
    }
  }

  private static void refreshPromotedMetadata(List<String> args) throws IOException {
    Path projectDirectory = projectDirectory(args);
    String targetKey = optionalValue(args, "--target");
    List<Path> metadataPaths;
    try (var stream = Files.walk(promotedMetadataRoot(projectDirectory))) {
      metadataPaths =
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .filter(path -> targetKey == null || path.getParent().getFileName().toString().equals(targetKey))
              .sorted()
              .toList();
    }

    for (Path metadataPath : metadataPaths) {
      PromotionMetadata metadata = JazzerJson.read(metadataPath, PromotionMetadata.class);
      JazzerRunTarget target = JazzerRunTarget.fromKey(metadata.targetKey());
      ReplayOutcome outcome =
          JazzerReplaySupport.replay(
              target.replayHarness(), Files.readAllBytes(Path.of(metadata.promotedInputPath())));
      Files.writeString(
          Path.of(metadata.replayTextPath()),
          JazzerTextRenderer.renderReplay(Path.of(metadata.promotedInputPath()), outcome)
              + System.lineSeparator());
      JazzerJson.write(
          metadataPath,
          new PromotionMetadata(
              metadata.targetKey(),
              metadata.sourcePath(),
              metadata.promotedInputPath(),
              JazzerReplaySupport.outcomeKind(outcome),
              JazzerReplaySupport.expectationFor(outcome),
              metadata.promotedAt(),
              metadata.replayTextPath()));
    }

    System.out.println("Refreshed " + metadataPaths.size() + " promoted metadata entries.");
  }

  private static List<LocalRunSummary> availableSummaries(Path projectDirectory) throws IOException {
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
    String configured = optionalValue(args, "--project-dir");
    return configured == null ? Path.of("").toAbsolutePath().normalize() : Path.of(configured);
  }

  private static Path promotedMetadataRoot(Path projectDirectory) {
    return projectDirectory.resolve("src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata");
  }

  private static Path requiredPath(List<String> args, String flag) {
    return Path.of(requiredValue(args, flag)).toAbsolutePath().normalize();
  }

  private static String requiredValue(List<String> args, String flag) {
    String value = optionalValue(args, flag);
    if (value == null) {
      throw new IllegalArgumentException("Missing required option " + flag);
    }
    return value;
  }

  private static String optionalValue(List<String> args, String flag) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(flag, "flag must not be null");
    for (int index = 0; index < args.size() - 1; index++) {
      if (args.get(index).equals(flag)) {
        return args.get(index + 1);
      }
    }
    return null;
  }

  private static boolean hasFlag(List<String> args, String flag) {
    return args.contains(flag);
  }

  private static String extensionFor(Path inputPath) {
    String fileName = inputPath.getFileName().toString();
    int separator = fileName.lastIndexOf('.');
    return separator >= 0 ? fileName.substring(separator) : ".bin";
  }

}
