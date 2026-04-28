package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.nio.file.Path;
import java.util.Objects;

/** Request path, source, and persistence facts shared across executor workflows. */
final class ExecutionRequestPaths {
  private ExecutionRequestPaths() {}

  static ExcelOoxmlPersistenceOptions persistenceOptions(
      WorkbookPlan.WorkbookPersistence persistence) {
    return persistenceOptions(persistence, Path.of(""));
  }

  static ExcelOoxmlPersistenceOptions persistenceOptions(
      WorkbookPlan.WorkbookPersistence persistence, Path workingDirectory) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> new ExcelOoxmlPersistenceOptions(null, null);
      case WorkbookPlan.WorkbookPersistence.OverwriteSource overwrite ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(
              overwrite.security().orElse(null), workingDirectory);
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(
              saveAs.security().orElse(null), workingDirectory);
    };
  }

  static ExcelOoxmlPackageSecuritySnapshot sourcePackageSecurity(
      WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> ExcelOoxmlPackageSecuritySnapshot.none();
      case WorkbookPlan.WorkbookSource.ExistingFile _ -> ExcelOoxmlPackageSecuritySnapshot.none();
    };
  }

  static String sourceEncryptionPassword(WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> null;
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          existingFile.security().map(security -> security.password()).orElse(null);
    };
  }

  static String persistencePath(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    return persistencePath(source, persistence, Path.of(""));
  }

  static String persistencePath(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory) {
    return normalizedPersistencePath(source, persistence, workingDirectory);
  }

  static WorkbookLocation workbookLocationFor(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    return workbookLocationFor(source, persistence, Path.of(""));
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
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  static String reqSourcePath(WorkbookPlan request) {
    return reqSourcePath(request, Path.of(""));
  }

  static String reqSourcePath(WorkbookPlan request, Path workingDirectory) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> null;
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          normalizePath(existingFile.path(), workingDirectory).toString();
    };
  }

  static dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape requestShape(
      WorkbookPlan request) {
    return dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known(
        reqSourceType(request), reqPersistenceType(request));
  }

  static dev.erst.gridgrind.contract.dto.ProblemContext.WorkbookReference workbookReference(
      WorkbookPlan request, Path workingDirectory) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ ->
          dev.erst.gridgrind.contract.dto.ProblemContext.WorkbookReference.newWorkbook();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          dev.erst.gridgrind.contract.dto.ProblemContext.WorkbookReference.existingFile(
              normalizePath(existingFile.path(), workingDirectory).toString());
    };
  }

  static dev.erst.gridgrind.contract.dto.ProblemContext.PersistenceReference persistenceReference(
      WorkbookPlan request, Path workingDirectory) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ ->
          dev.erst.gridgrind.contract.dto.ProblemContext.PersistenceReference.overwriteSource(
              Objects.requireNonNull(reqSourcePath(request, workingDirectory)));
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          dev.erst.gridgrind.contract.dto.ProblemContext.PersistenceReference.saveAs(
              normalizePath(saveAs.path(), workingDirectory).toString());
      case WorkbookPlan.WorkbookPersistence.None _ ->
          throw new IllegalArgumentException("persistence reference requires a saving policy");
    };
  }

  static Path normalizePath(String path) {
    return normalizePath(path, Path.of(""));
  }

  static Path normalizePath(String path, Path workingDirectory) {
    Path candidate = Path.of(path);
    if (candidate.isAbsolute()) {
      return candidate.toAbsolutePath().normalize();
    }
    return workingDirectory.toAbsolutePath().normalize().resolve(candidate).normalize();
  }

  private static String normalizedPersistencePath(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> null;
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          normalizePath(saveAs.path(), workingDirectory).toString();
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ ->
          source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile
              ? normalizePath(existingFile.path(), workingDirectory).toString()
              : null;
    };
  }
}
