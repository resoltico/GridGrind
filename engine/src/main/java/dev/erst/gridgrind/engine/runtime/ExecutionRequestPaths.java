package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureSpec;
import dev.erst.gridgrind.core.model.security.ResolvedEncryption;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceEncryption;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSignature;
import dev.erst.gridgrind.core.model.security.ResolvedSignature;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Request path, source, and persistence facts shared across executor workflows. */
final class ExecutionRequestPaths {
  private ExecutionRequestPaths() {}

  static ResolvedPersistenceSecurity persistenceOptions(
      WorkbookPlan.WorkbookPersistence persistence, ExecutionInputBindings bindings)
      throws IOException {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> ResolvedPersistenceSecurity.none();
      case WorkbookPlan.WorkbookPersistence.Overwrite overwrite ->
          materializePersistenceSecurity(overwrite.security().orElse(null), bindings);
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          materializePersistenceSecurity(saveAs.security().orElse(null), bindings);
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

  static OoxmlPackageSecurityReport sourcePackageSecurity(WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> OoxmlPackageSecurityReport.none();
      case WorkbookPlan.WorkbookSource.ExistingFile _ -> OoxmlPackageSecurityReport.none();
    };
  }

  static Optional<String> sourceEncryptionPassword(
      WorkbookPlan.WorkbookSource source, ExecutionInputBindings bindings) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> Optional.empty();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          existingFile
              .security()
              .flatMap(dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec::passwordRef)
              .map(bindings::resolveSecret);
    };
  }

  static Optional<String> sourceEncryptionPassword(WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> Optional.empty();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile -> {
        if (existingFile
            .security()
            .flatMap(dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec::passwordRef)
            .isPresent()) {
          throw new IllegalStateException(
              "source encryption references require execution bindings");
        }
        yield Optional.empty();
      }
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

  static dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape requestShape(
      WorkbookPlan request) {
    return dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape.known(
        reqSourceType(request), reqPersistenceType(request));
  }

  static dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.WorkbookReference
      workbookReference(WorkbookPlan request, Path workingDirectory) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ ->
          dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.WorkbookReference
              .newWorkbook();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.WorkbookReference
              .existingFile(normalizePath(existingFile.path(), workingDirectory).toString());
    };
  }

  static dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.PersistenceReference
      persistenceReference(WorkbookPlan request, Path workingDirectory) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.PersistenceReference
              .overwrite(Objects.requireNonNull(reqSourcePath(request, workingDirectory)));
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.PersistenceReference
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

  static ResolvedOpenSecurity materializeOpenSecurity(
      @Nullable OoxmlOpenSecuritySpec input, ExecutionInputBindings bindings) {
    return input == null || input.passwordRef().isEmpty()
        ? new ResolvedOpenSecurity.Unencrypted()
        : new ResolvedOpenSecurity.Encrypted(
            bindings.resolveSecret(input.passwordRef().orElseThrow()));
  }

  static ResolvedPersistenceSecurity materializePersistenceSecurity(
      @Nullable OoxmlPersistenceSecuritySpec input, ExecutionInputBindings bindings)
      throws IOException {
    if (input == null) {
      return ResolvedPersistenceSecurity.none();
    }
    return new ResolvedPersistenceSecurity(
        encryptionPolicy(input.encryption(), bindings),
        signaturePolicy(input.signature(), bindings));
  }

  private static ResolvedPersistenceEncryption encryptionPolicy(
      OoxmlPersistenceEncryptionSpec input, ExecutionInputBindings bindings) {
    return switch (input) {
      case OoxmlPersistenceEncryptionSpec.None none -> none;
      case OoxmlPersistenceEncryptionSpec.Encrypt encrypt ->
          new ResolvedPersistenceEncryption.Encrypt(
              materializeEncryption(encrypt.encryption(), bindings));
      case OoxmlPersistenceEncryptionSpec.PreserveSource preserve -> preserve;
    };
  }

  private static ResolvedPersistenceSignature signaturePolicy(
      OoxmlPersistenceSignatureSpec input, ExecutionInputBindings bindings) throws IOException {
    return switch (input) {
      case OoxmlPersistenceSignatureSpec.None none -> none;
      case OoxmlPersistenceSignatureSpec.Sign sign ->
          new ResolvedPersistenceSignature.Sign(materializeSignature(sign.signature(), bindings));
    };
  }

  private static ResolvedEncryption materializeEncryption(
      OoxmlEncryptionSpec input, ExecutionInputBindings bindings) {
    return new ResolvedEncryption(
        bindings.resolveSecret(input.passwordRef()), input.cipher(), input.hash());
  }

  private static ResolvedSignature materializeSignature(
      OoxmlSignatureSpec input, ExecutionInputBindings bindings) throws IOException {
    try {
      return new ResolvedSignature(
          bindings
              .requestPathAccess()
              .materializeRead(
                  input.pkcs12Path(),
                  "persistence.security.signature.signature.pkcs12Path",
                  "gridgrind-signing-material-",
                  ".p12"),
          bindings.resolveSecret(input.keystorePasswordRef()),
          input
              .keyPasswordRef()
              .map(bindings::resolveSecret)
              .orElseGet(() -> bindings.resolveSecret(input.keystorePasswordRef())),
          input.alias(),
          input.digestAlgorithm(),
          input.description());
    } catch (java.nio.file.NoSuchFileException exception) {
      throw new dev.erst.gridgrind.engine.backend.xssf.InvalidSigningConfigurationException(
          "Signing material does not exist: " + input.pkcs12Path(), exception);
    }
  }
}
