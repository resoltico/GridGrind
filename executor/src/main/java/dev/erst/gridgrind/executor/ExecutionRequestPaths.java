package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.nio.file.Path;

/** Request path, source, and persistence facts shared across executor workflows. */
final class ExecutionRequestPaths {
  private ExecutionRequestPaths() {}

  static ExcelOoxmlPersistenceOptions persistenceOptions(
      WorkbookPlan.WorkbookPersistence persistence) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> new ExcelOoxmlPersistenceOptions(null, null);
      case WorkbookPlan.WorkbookPersistence.OverwriteSource overwrite ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(overwrite.security());
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(saveAs.security());
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
          existingFile.security() == null ? null : existingFile.security().password();
    };
  }

  static String persistencePath(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    return normalizedPersistencePath(source, persistence);
  }

  static WorkbookLocation workbookLocationFor(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    String persistencePath = normalizedPersistencePath(source, persistence);
    if (persistencePath != null) {
      return new WorkbookLocation.StoredWorkbook(Path.of(persistencePath));
    }
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> new WorkbookLocation.UnsavedWorkbook();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          new WorkbookLocation.StoredWorkbook(normalizePath(existingFile.path()));
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
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> null;
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          normalizePath(existingFile.path()).toString();
    };
  }

  static Path normalizePath(String path) {
    return Path.of(path).toAbsolutePath().normalize();
  }

  private static String normalizedPersistencePath(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> null;
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          normalizePath(saveAs.path()).toString();
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ ->
          source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile
              ? normalizePath(existingFile.path()).toString()
              : null;
    };
  }
}
