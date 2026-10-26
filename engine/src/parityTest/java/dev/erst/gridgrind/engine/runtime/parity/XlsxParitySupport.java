package dev.erst.gridgrind.engine.runtime.parity;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.engine.runtime.ExecutionInputBindings;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.Callable;

/** Shared exception-wrapping support for the parity ledger, corpus, and oracle harness. */
final class XlsxParitySupport {
  private static final Path MANAGED_TEMP_SEGMENT = Path.of(".gridgrind", "tmp");
  private static final java.util.Set<String> CORPUS_ROOT_CHILDREN =
      java.util.Set.of("corpus", "scenario-copies", "workbooks");

  private XlsxParitySupport() {}

  static ExecutionInputBindings bindings(Path executionRoot, WorkbookPlan request) {
    Path normalizedRoot = executionRoot.toAbsolutePath().normalize();
    Map<SecretReference, String> secrets = fixtureSecrets(request);
    return new ExecutionInputBindings(
        normalizedRoot,
        managedTempRoot(normalizedRoot),
        new ExecutionGrant(
            readableResources(request, normalizedRoot),
            request.steps().stream().map(XlsxParitySupport::operationId).distinct().toList(),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            publicationAuthority(request.persistence(), normalizedRoot),
            secrets.keySet().stream().toList(),
            dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum()),
        reference -> secrets.get(reference).toCharArray());
  }

  static WorkbookTempFileFactory tempFileFactory(Path executionRoot) {
    return WorkbookTempFileFactory.rooted(managedTempRoot(executionRoot));
  }

  static Path executionRootFor(Path anchoredPath) {
    Path normalized = anchoredPath.toAbsolutePath().normalize();
    Path parent = normalized.getParent();
    for (Path current = parent; current != null; current = current.getParent()) {
      Path name = current.getFileName();
      if (name != null && CORPUS_ROOT_CHILDREN.contains(name.toString())) {
        Path corpusRoot = current.getParent();
        return corpusRoot == null ? current : corpusRoot;
      }
    }
    return parent == null ? normalized : parent;
  }

  private static Path managedTempRoot(Path executionRoot) {
    return executionRoot.toAbsolutePath().normalize().resolve(MANAGED_TEMP_SEGMENT);
  }

  private static java.util.List<ExecutionGrant.ReadAuthority> readableResources(
      WorkbookPlan request, Path executionRoot) {
    java.util.List<Path> paths = new java.util.ArrayList<>();
    collectDeclaredPaths(GridGrindJsonOutput.requestTree(request), executionRoot, paths);
    return paths.stream()
        .distinct()
        .<ExecutionGrant.ReadAuthority>map(ExecutionGrant.ReadAuthority.File::new)
        .toList();
  }

  private static void collectDeclaredPaths(
      tools.jackson.databind.JsonNode node, Path executionRoot, java.util.List<Path> paths) {
    switch (node) {
      case tools.jackson.databind.node.ObjectNode object -> {
        for (var property : object.properties()) {
          if (property.getKey().toLowerCase(java.util.Locale.ROOT).endsWith("path")
              && property.getValue().isString()) {
            paths.add(resolve(property.getValue().asString(), executionRoot));
          } else {
            collectDeclaredPaths(property.getValue(), executionRoot, paths);
          }
        }
      }
      case tools.jackson.databind.node.ArrayNode array -> {
        for (tools.jackson.databind.JsonNode element : array) {
          collectDeclaredPaths(element, executionRoot, paths);
        }
      }
      default -> {}
    }
  }

  private static Map<SecretReference, String> fixtureSecrets(WorkbookPlan request) {
    Map<SecretReference, String> known =
        Map.of(
            new SecretReference("source-password"), XlsxParityScenarios.ENCRYPTION_PASSWORD,
            new SecretReference("output-password"), XlsxParityScenarios.ENCRYPTION_PASSWORD,
            new SecretReference("keystore-password"), XlsxParityScenarios.SIGNING_KEYSTORE_PASSWORD,
            new SecretReference("key-password"), XlsxParityScenarios.SIGNING_KEY_PASSWORD,
            new SecretReference("workbook-password"),
                XlsxParityScenarios.WORKBOOK_PROTECTION_PASSWORD,
            new SecretReference("sheet-password"), XlsxParityScenarios.SHEET_PROTECTION_PASSWORD,
            new SecretReference("wrong-password"), "gridgrind-phase9-wrong-password");
    List<SecretReference> references = new ArrayList<>();
    collectSecretReferences(GridGrindJsonOutput.requestTree(request), references);
    return known.entrySet().stream()
        .filter(entry -> references.contains(entry.getKey()))
        .collect(java.util.stream.Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));
  }

  private static void collectSecretReferences(
      tools.jackson.databind.JsonNode node, List<SecretReference> references) {
    switch (node) {
      case tools.jackson.databind.node.ObjectNode object -> {
        for (var property : object.properties()) {
          if (property.getKey().endsWith("Ref")
              && property.getValue() instanceof tools.jackson.databind.node.ObjectNode ref
              && ref.path("id").isString()) {
            references.add(new SecretReference(ref.path("id").asString()));
          } else {
            collectSecretReferences(property.getValue(), references);
          }
        }
      }
      case tools.jackson.databind.node.ArrayNode array -> {
        for (tools.jackson.databind.JsonNode element : array) {
          collectSecretReferences(element, references);
        }
      }
      default -> {}
    }
  }

  private static String operationId(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutation -> mutation.action().actionType();
      case AssertionStep assertion -> assertion.assertion().assertionType();
      case InspectionStep inspection -> inspection.query().queryType();
    };
  }

  private static ExecutionGrant.PublicationAuthority publicationAuthority(
      WorkbookPlan.WorkbookPersistence persistence, Path executionRoot) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new ExecutionGrant.PublicationAuthority.None();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          new ExecutionGrant.PublicationAuthority.SaveAs(
              resolve(saveAs.path(), executionRoot), saveAs.ifExists());
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          new ExecutionGrant.PublicationAuthority.OverwriteSource();
    };
  }

  private static Path resolve(String rawPath, Path executionRoot) {
    Path candidate = Path.of(rawPath);
    return (candidate.isAbsolute() ? candidate : executionRoot.resolve(candidate))
        .toAbsolutePath()
        .normalize();
  }

  static <T> T call(String action, Callable<T> callable) {
    Objects.requireNonNull(action, "action must not be null");
    Objects.requireNonNull(callable, "callable must not be null");
    try {
      return callable.call();
    } catch (RuntimeException exception) {
      throw exception;
    } catch (Exception exception) {
      throw new XlsxParityException(action, exception);
    }
  }
}

/** Runtime failure wrapper used when parity harness work cannot complete successfully. */
final class XlsxParityException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  XlsxParityException(String action, Exception cause) {
    super(action, cause);
  }
}
