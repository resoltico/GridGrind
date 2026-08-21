package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Request path, source, and persistence facts shared across executor workflows. */
final class ExecutionRequestPaths {
  private ExecutionRequestPaths() {}

  static ExcelOoxmlPersistenceOptions persistenceOptions(
      WorkbookPlan.WorkbookPersistence persistence, ExecutionInputBindings bindings)
      throws IOException {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> ExcelOoxmlPersistenceOptions.none();
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(
              overwrite.security().orElse(null), bindings);
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(
              saveAs.security().orElse(null), bindings);
    };
  }

  static WorkbookArtifactWriteDisposition writeDisposition(
      WorkbookPlan.WorkbookPersistence.SaveAs saveAs) {
    Objects.requireNonNull(saveAs, "saveAs must not be null");
    return switch (saveAs.ifExists()) {
      case REJECT -> WorkbookArtifactWriteDisposition.CREATE_NEW;
      case REPLACE -> WorkbookArtifactWriteDisposition.REPLACE_EXISTING;
    };
  }

  static ExcelOoxmlPackageSecuritySnapshot sourcePackageSecurity(
      WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> ExcelOoxmlPackageSecuritySnapshot.none();
      case WorkbookPlan.WorkbookSource.ExistingFile _ -> ExcelOoxmlPackageSecuritySnapshot.none();
    };
  }

  static Optional<String> sourceEncryptionPassword(WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> Optional.empty();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          existingFile
              .security()
              .flatMap(dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput::password);
    };
  }

  static @org.jspecify.annotations.Nullable String persistencePath(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory) {
    return normalizedPersistencePath(source, persistence, workingDirectory);
  }

  static WorkbookLocation workbookLocationFor(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory) {
    String persistencePath = normalizedPersistencePath(source, persistence, workingDirectory);
    if (persistencePath != null) {
      return new WorkbookLocation.StoredWorkbook(Path.of(persistencePath));
    }
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> new WorkbookLocation.UnsavedWorkbook();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          new WorkbookLocation.StoredWorkbook(normalizePath(existingFile.path(), workingDirectory));
    };
  }

  static String reqSourceType(WorkbookPlan request) {
    if (request.source() instanceof WorkbookPlan.WorkbookSource.New) {
      return "NEW";
    }
    return "EXISTING";
  }

  static String reqPersistenceType(WorkbookPlan request) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> "OVERWRITE";
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  static @Nullable String reqSourcePath(WorkbookPlan request, Path workingDirectory) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> null;
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          normalizePath(existingFile.path(), workingDirectory).toString();
    };
  }

  static dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape requestShape(
      WorkbookPlan request) {
    return dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
        reqSourceType(request), reqPersistenceType(request));
  }

  static dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.WorkbookReference
      workbookReference(WorkbookPlan request, Path workingDirectory) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ ->
          dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.WorkbookReference
              .newWorkbook();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.WorkbookReference
              .existingFile(normalizePath(existingFile.path(), workingDirectory).toString());
    };
  }

  static dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.PersistenceReference
      persistenceReference(WorkbookPlan request, Path workingDirectory) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.PersistenceReference
              .overwrite(Objects.requireNonNull(reqSourcePath(request, workingDirectory)));
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.PersistenceReference
              .saveAs(normalizePath(saveAs.path(), workingDirectory).toString());
      case WorkbookPlan.WorkbookPersistence.None _ ->
          throw new IllegalArgumentException("persistence reference requires a saving policy");
    };
  }

  static Path normalizePath(String path, Path workingDirectory) {
    Path candidate = Path.of(path);
    Path base = workingDirectory.toAbsolutePath().normalize();
    Path normalized =
        candidate.isAbsolute()
            ? candidate.toAbsolutePath().normalize()
            : base.resolve(candidate).normalize();
    if (!normalized.startsWith(base)) { // LIM-025
      throw new RequestPathEscapeException("path must not escape the execution root: " + path);
    }
    return normalized;
  }

  private static @Nullable String normalizedPersistencePath(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> null;
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          normalizePath(saveAs.path(), workingDirectory).toString();
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile
              ? normalizePath(existingFile.path(), workingDirectory).toString()
              : null;
    };
  }
}
