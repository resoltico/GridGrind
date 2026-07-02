package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import java.util.Map;
import java.util.Objects;

/** Canonical English term surfaces used by deterministic task keyword matching. */
final class GridGrindTaskKeywordSurfaces {
  private static final Map<TaskSourceMode, String> SOURCE_MODE_SURFACES =
      Map.of(
          TaskSourceMode.NEW_WORKBOOK, "new blank create workbook",
          TaskSourceMode.EXISTING_WORKBOOK, "existing open input workbook");
  private static final Map<TaskPersistenceMode, String> PERSISTENCE_MODE_SURFACES =
      Map.of(
          TaskPersistenceMode.NONE, "no save transient non destructive",
          TaskPersistenceMode.SAVE_AS, "save export output",
          TaskPersistenceMode.OVERWRITE, "overwrite in place");
  private static final Map<TaskMutationMode, String> MUTATION_MODE_SURFACES =
      Map.of(
          TaskMutationMode.READ_ONLY, "read only inspect audit analyze",
          TaskMutationMode.MUTATING, "author mutate build update");
  private static final Map<TaskAssetMode, String> ASSET_MODE_SURFACES =
      Map.of(
          TaskAssetMode.SELF_CONTAINED, "self contained portable",
          TaskAssetMode.REQUIRES_EXTERNAL_PAYLOADS, "external payload file assets");
  private static final Map<TaskPhasePurpose, String> PHASE_PURPOSE_SURFACES =
      Map.of(
          TaskPhasePurpose.PREPARE, "prepare setup scaffold",
          TaskPhasePurpose.AUTHOR, "author build mutate",
          TaskPhasePurpose.IMPORT, "import ingest",
          TaskPhasePurpose.EXPORT, "export extract",
          TaskPhasePurpose.INSPECT, "inspect readback",
          TaskPhasePurpose.ANALYZE, "analyze findings health",
          TaskPhasePurpose.VERIFY, "verify assert confirm");
  private static final Map<TaskInputKind, String> INPUT_KIND_SURFACES =
      Map.ofEntries(
          Map.entry(TaskInputKind.SOURCE_WORKBOOK_PATH, "existing workbook input source file path"),
          Map.entry(TaskInputKind.PERSISTENCE_TARGET_PATH, "save export output destination path"),
          Map.entry(TaskInputKind.TARGET_SHEET_NAMES, "sheet name worksheet names"),
          Map.entry(
              TaskInputKind.TARGET_OBJECT_NAMES,
              "table pivot chart named range drawing object names"),
          Map.entry(
              TaskInputKind.CELL_OR_RANGE_COORDINATES,
              "cell range coordinates address anchor placement"),
          Map.entry(TaskInputKind.TABULAR_SOURCE_ROWS, "source rows table data values"),
          Map.entry(
              TaskInputKind.VALIDATION_RULES, "validation rules allowed values protection prompts"),
          Map.entry(TaskInputKind.MAPPING_LOCATOR, "mapping locator map id mapping name"),
          Map.entry(TaskInputKind.XML_PAYLOAD, "xml payload mapping document import export"),
          Map.entry(TaskInputKind.BINARY_PAYLOAD, "binary payload image object file"),
          Map.entry(TaskInputKind.DRAWING_ANCHORS, "drawing anchor position placement"));
  private static final Map<TaskGoalKind, String> GOAL_SURFACES =
      Map.of(
          TaskGoalKind.AUTHOR, "author create build",
          TaskGoalKind.VERIFY, "verify validate confirm",
          TaskGoalKind.INSPECT, "inspect read discover",
          TaskGoalKind.ANALYZE, "analyze findings diagnose",
          TaskGoalKind.IMPORT, "import ingest load",
          TaskGoalKind.EXPORT, "export extract emit",
          TaskGoalKind.MAINTAIN, "maintain normalize repair");
  private static final Map<TaskArtifactKind, String> ARTIFACT_SURFACES =
      Map.ofEntries(
          Map.entry(TaskArtifactKind.WORKBOOK, "workbook spreadsheet file"),
          Map.entry(TaskArtifactKind.SHEET, "sheet worksheet tab"),
          Map.entry(TaskArtifactKind.CELL, "cell grid value"),
          Map.entry(TaskArtifactKind.TABLE, "table rows columns"),
          Map.entry(TaskArtifactKind.CHART, "chart graph dashboard"),
          Map.entry(TaskArtifactKind.PIVOT_TABLE, "pivot pivot table summary"),
          Map.entry(TaskArtifactKind.DATA_VALIDATION, "validation input rules"),
          Map.entry(TaskArtifactKind.COMMENT, "comment annotation remark"),
          Map.entry(TaskArtifactKind.PROTECTION, "protection locked guardrail"),
          Map.entry(TaskArtifactKind.CUSTOM_XML_MAPPING, "custom xml mapping"),
          Map.entry(TaskArtifactKind.XML_PAYLOAD, "xml payload import export"),
          Map.entry(TaskArtifactKind.DRAWING_OBJECT, "drawing picture shape object"),
          Map.entry(TaskArtifactKind.SIGNATURE_LINE, "signature line signing"),
          Map.entry(TaskArtifactKind.PACKAGE_SECURITY, "package security encryption ooxml"),
          Map.entry(TaskArtifactKind.FORMULA_SURFACE, "formula surface formulas"),
          Map.entry(TaskArtifactKind.NAMED_RANGE, "named range name reference"));
  private static final Map<TaskVerificationKind, String> VERIFICATION_KIND_SURFACES =
      Map.of(
          TaskVerificationKind.FACT_READBACK, "inspect readback factual verification",
          TaskVerificationKind.ASSERTION_CHECKS, "assert assertion invariant checks",
          TaskVerificationKind.HEALTH_ANALYSIS, "health analysis findings audit",
          TaskVerificationKind.EXPORT_REREAD, "export reread roundtrip verification");

  private GridGrindTaskKeywordSurfaces() {}

  static String sourceModeSurface(TaskSourceMode sourceMode) {
    return requiredSurface(SOURCE_MODE_SURFACES, sourceMode, "source mode");
  }

  static String persistenceModeSurface(TaskPersistenceMode persistenceMode) {
    return requiredSurface(PERSISTENCE_MODE_SURFACES, persistenceMode, "persistence mode");
  }

  static String mutationModeSurface(TaskMutationMode mutationMode) {
    return requiredSurface(MUTATION_MODE_SURFACES, mutationMode, "mutation mode");
  }

  static String assetModeSurface(TaskAssetMode assetMode) {
    return requiredSurface(ASSET_MODE_SURFACES, assetMode, "asset mode");
  }

  static String phasePurposeSurface(TaskPhasePurpose purpose) {
    return requiredSurface(PHASE_PURPOSE_SURFACES, purpose, "phase purpose");
  }

  static String inputKindSurface(TaskInputKind inputKind) {
    return requiredSurface(INPUT_KIND_SURFACES, inputKind, "input kind");
  }

  static String goalSurface(TaskGoalKind goal) {
    return requiredSurface(GOAL_SURFACES, goal, "goal");
  }

  static String artifactSurface(TaskArtifactKind artifact) {
    return requiredSurface(ARTIFACT_SURFACES, artifact, "artifact");
  }

  static String verificationKindSurface(TaskVerificationKind verificationKind) {
    return requiredSurface(VERIFICATION_KIND_SURFACES, verificationKind, "verification kind");
  }

  private static <K> String requiredSurface(Map<K, String> surfaces, K key, String surfaceLabel) {
    Objects.requireNonNull(key, surfaceLabel + " must not be null");
    return Objects.requireNonNull(surfaces.get(key), "Unsupported " + surfaceLabel + ": " + key);
  }
}
