package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.excel.OoxmlSecurityTestSupport;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.ObjectNode;

/** Test helper for explicit execution-root bindings with one managed scratch root. */
final class ExecutionInputBindingsFixtureSupport {
  private static final Path MANAGED_TEMP_SEGMENT = Path.of(".gridgrind", "tmp");

  private ExecutionInputBindingsFixtureSupport() {}

  static ExecutionInputBindings bindings(Path workingDirectory) {
    return new ExecutionInputBindings(
        workingDirectory, managedTempRoot(workingDirectory), noPublicationGrant());
  }

  static ExecutionInputBindings bindings(Path workingDirectory, byte[] standardInputBytes) {
    return new ExecutionInputBindings(
        workingDirectory,
        managedTempRoot(workingDirectory),
        standardInputBytes,
        standardInputGrant());
  }

  static ExecutionInputBindings bindings(Path workingDirectory, WorkbookPlan plan) {
    Map<SecretReference, String> secrets = deterministicFixtureSecrets(plan);
    if (!secrets.isEmpty()) {
      return bindings(workingDirectory, plan, secrets);
    }
    return new ExecutionInputBindings(
        workingDirectory, managedTempRoot(workingDirectory), grantForPlan(plan, workingDirectory));
  }

  static ExecutionInputBindings bindings(
      Path workingDirectory,
      WorkbookPlan plan,
      dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy hostAcceptancePolicy) {
    GridGrindExecutionGrant.Bounded planGrant = grantForPlan(plan, workingDirectory);
    return new ExecutionInputBindings(
        workingDirectory,
        managedTempRoot(workingDirectory),
        new GridGrindExecutionGrant.Bounded(
            planGrant.readableResources(),
            planGrant.operationIds(),
            planGrant.workbookTargetAuthority(),
            planGrant.publicationAuthority(),
            planGrant.allowedSecretReferences(),
            java.util.Objects.requireNonNull(
                hostAcceptancePolicy, "hostAcceptancePolicy must not be null")));
  }

  static ExecutionInputBindings bindings(
      Path workingDirectory, WorkbookPlan plan, Map<SecretReference, String> secrets) {
    GridGrindExecutionGrant.Bounded planGrant = grantForPlan(plan, workingDirectory);
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            planGrant.readableResources(),
            planGrant.operationIds(),
            planGrant.workbookTargetAuthority(),
            planGrant.publicationAuthority(),
            secrets.keySet().stream().toList(),
            planGrant.hostAcceptancePolicy());
    return new ExecutionInputBindings(
        workingDirectory,
        managedTempRoot(workingDirectory),
        grant,
        reference -> {
          String value = secrets.get(reference);
          if (value == null) {
            throw new IllegalArgumentException(
                "test host has no value for secret reference " + reference.id());
          }
          return value.toCharArray();
        });
  }

  static ExecutionInputBindings bindings(
      Path workingDirectory, WorkbookPlan plan, byte[] standardInputBytes) {
    GridGrindExecutionGrant.Bounded planGrant = grantForPlan(plan, workingDirectory);
    List<GridGrindExecutionGrant.ReadAuthority> resources =
        java.util.stream.Stream.concat(
                planGrant.readableResources().stream(),
                java.util.stream.Stream.of(
                    new GridGrindExecutionGrant.ReadAuthority.StandardInput()))
            .toList();
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            resources,
            planGrant.operationIds(),
            planGrant.workbookTargetAuthority(),
            planGrant.publicationAuthority(),
            planGrant.allowedSecretReferences(),
            planGrant.hostAcceptancePolicy());
    return new ExecutionInputBindings(
        workingDirectory, managedTempRoot(workingDirectory), standardInputBytes, grant);
  }

  static ExecutionInputBindings bindings(Path workingDirectory, List<Path> readableFiles) {
    return new ExecutionInputBindings(
        workingDirectory,
        managedTempRoot(workingDirectory),
        noPublicationGrant(List.of(), readableFiles));
  }

  static PreparedBindings preparedBindings(Path workingDirectory) {
    ExecutionInputBindings bindings = bindings(workingDirectory);
    return preparedBindings(workingDirectory, bindings);
  }

  static PreparedBindings preparedBindings(Path workingDirectory, List<Path> readableFiles) {
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            workingDirectory,
            managedTempRoot(workingDirectory),
            noPublicationGrant(List.of(), readableFiles));
    return preparedBindings(workingDirectory, bindings);
  }

  static PreparedBindings preparedBindings(
      Path workingDirectory, List<Path> readableFiles, Map<SecretReference, String> secrets) {
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            readableFiles.stream()
                .<GridGrindExecutionGrant.ReadAuthority>map(
                    GridGrindExecutionGrant.ReadAuthority.File::new)
                .toList(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            secrets.keySet().stream().toList(),
            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            workingDirectory,
            managedTempRoot(workingDirectory),
            grant,
            reference -> {
              String value = secrets.get(reference);
              if (value == null) {
                throw new IllegalArgumentException(
                    "test host has no value for secret reference " + reference.id());
              }
              return value.toCharArray();
            });
    return preparedBindings(workingDirectory, bindings);
  }

  private static PreparedBindings preparedBindings(
      Path workingDirectory, ExecutionInputBindings bindings) {
    RequestPathAccess access =
        new RequestPathAccess(
            workingDirectory, bindings.tempFileFactory(), bindings.executionGrant());
    return new PreparedBindings(bindings.withRequestPathAccess(access), access);
  }

  private static Path managedTempRoot(Path workingDirectory) {
    return workingDirectory.toAbsolutePath().normalize().resolve(MANAGED_TEMP_SEGMENT);
  }

  static GridGrindExecutionGrant.Bounded noPublicationGrant() {
    return noPublicationGrant(List.of(), List.of());
  }

  static GridGrindExecutionGrant.Bounded noPublicationGrant(
      List<String> operationIds, List<Path> readableFiles) {
    return new GridGrindExecutionGrant.Bounded(
        readableFiles.stream()
            .<GridGrindExecutionGrant.ReadAuthority>map(
                GridGrindExecutionGrant.ReadAuthority.File::new)
            .toList(),
        operationIds,
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new GridGrindExecutionGrant.PublicationAuthority.None(),
        List.of(),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }

  static GridGrindExecutionGrant.Bounded grantForPlan(WorkbookPlan plan, Path workingDirectory) {
    Path normalizedWorkingDirectory = workingDirectory.toAbsolutePath().normalize();
    List<GridGrindExecutionGrant.ReadAuthority> readableResources =
        declaredReadableFiles(plan, normalizedWorkingDirectory).stream()
            .<GridGrindExecutionGrant.ReadAuthority>map(
                GridGrindExecutionGrant.ReadAuthority.File::new)
            .toList();
    return new GridGrindExecutionGrant.Bounded(
        readableResources,
        plan.steps().stream().map(ExecutionInputBindingPlanFacts::operationId).distinct().toList(),
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        ExecutionInputBindingPlanFacts.publicationAuthority(
            plan.persistence(), normalizedWorkingDirectory),
        declaredSecretReferences(plan),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }

  private static List<SecretReference> declaredSecretReferences(WorkbookPlan plan) {
    List<SecretReference> references = new ArrayList<>();
    if (plan.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile) {
      existingFile
          .security()
          .flatMap(dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput::passwordRef)
          .ifPresent(references::add);
    }
    java.util.Optional<dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput> security =
        switch (plan.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> java.util.Optional.empty();
          case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> saveAs.security();
          case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> overwrite.security();
        };
    security.ifPresent(
        value -> {
          if (value.encryption()
              instanceof
              dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.Encrypt encrypt) {
            references.add(encrypt.encryption().passwordRef());
          }
          if (value.signature()
              instanceof dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.Sign sign) {
            references.add(sign.signature().keystorePasswordRef());
            sign.signature().keyPasswordRef().ifPresent(references::add);
          }
        });
    return references.stream().distinct().toList();
  }

  private static Map<SecretReference, String> deterministicFixtureSecrets(WorkbookPlan plan) {
    @SuppressWarnings("PMD.UseConcurrentHashMap")
    Map<SecretReference, String> values = new java.util.LinkedHashMap<>();
    for (SecretReference reference : declaredSecretReferences(plan)) {
      String value =
          switch (reference.id()) {
            case "source-password", "output-password" ->
                OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD;
            case "keystore-password" -> OoxmlSecurityTestSupport.KEYSTORE_PASSWORD;
            case "key-password" -> OoxmlSecurityTestSupport.KEY_PASSWORD;
            case "wrong-password" -> "wrong-password";
            default -> null;
          };
      if (value != null) {
        values.put(reference, value);
      }
    }
    return Map.copyOf(values);
  }

  private static GridGrindExecutionGrant.Bounded standardInputGrant() {
    return new GridGrindExecutionGrant.Bounded(
        List.of(new GridGrindExecutionGrant.ReadAuthority.StandardInput()),
        List.of(),
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new GridGrindExecutionGrant.PublicationAuthority.None(),
        List.of(),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }

  private static List<Path> declaredReadableFiles(WorkbookPlan plan, Path workingDirectory) {
    List<Path> paths = new ArrayList<>();
    try {
      collectDeclaredPaths(GridGrindJsonOutput.requestTree(plan), workingDirectory, paths);
    } catch (RuntimeException ignored) {
      switch (plan.source()) {
        case WorkbookPlan.WorkbookSource.New _ -> {}
        case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
            paths.add(resolve(existingFile.path(), workingDirectory));
      }
      plan.formulaEnvironment().externalWorkbooks().stream()
          .map(workbook -> resolve(workbook.path(), workingDirectory))
          .forEach(paths::add);
    }
    return paths.stream().distinct().toList();
  }

  private static void collectDeclaredPaths(JsonNode node, Path workingDirectory, List<Path> paths) {
    switch (node) {
      case ObjectNode object -> {
        for (var property : object.properties()) {
          if (property.getKey().toLowerCase(java.util.Locale.ROOT).endsWith("path")
              && property.getValue().isString()) {
            paths.add(resolve(property.getValue().asString(), workingDirectory));
          } else {
            collectDeclaredPaths(property.getValue(), workingDirectory, paths);
          }
        }
      }
      case ArrayNode array -> {
        for (JsonNode element : array) {
          collectDeclaredPaths(element, workingDirectory, paths);
        }
      }
      default -> {}
    }
  }

  private static Path resolve(String rawPath, Path workingDirectory) {
    Path candidate = Path.of(rawPath);
    return (candidate.isAbsolute() ? candidate : workingDirectory.resolve(candidate))
        .toAbsolutePath()
        .normalize();
  }

  record PreparedBindings(ExecutionInputBindings bindings, RequestPathAccess access)
      implements AutoCloseable {
    PreparedBindings {
      Objects.requireNonNull(bindings, "bindings must not be null");
      Objects.requireNonNull(access, "access must not be null");
    }

    @Override
    public void close() throws IOException {
      access.close();
    }
  }
}
