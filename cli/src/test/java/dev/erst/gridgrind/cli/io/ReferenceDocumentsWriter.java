package dev.erst.gridgrind.cli.io;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.core.catalog.TypeEntry;
import dev.erst.gridgrind.core.limits.EnforcesLimit;
import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport;
import dev.erst.gridgrind.json.schema.JsonSchemas;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Produces reference pages from the same metadata consumed by the wire catalog. */
public final class ReferenceDocumentsWriter {
  private static final String GENERATED = "<!-- Generated from core metadata; do not edit. -->";
  private static final String FENCE = Character.toString((char) 96).repeat(3);

  private ReferenceDocumentsWriter() {}

  public static void main(String[] args) throws IOException {
    if (args.length != 2 || !List.of("--write", "--verify").contains(args[1])) {
      throw new IllegalArgumentException("ReferenceDocumentsWriter <repository> --write|--verify");
    }
    Path root = Path.of(args[0]).toAbsolutePath().normalize();
    boolean verify = "--verify".equals(args[1]);
    Map<Path, String> documents = render(root);
    if (verify) {
      for (JsonSchemas.Document schema : JsonSchemas.Document.values()) {
        documents.put(
            root.resolve("docs/reference/schema").resolve(schema.fileName()),
            JsonSchemas.document(schema));
      }
    }
    for (var entry : documents.entrySet()) publish(entry.getKey(), entry.getValue(), verify);
    removeObsoletePages(root, documents, verify);
  }

  private static void publish(Path file, String expected, boolean verify) throws IOException {
    if (verify) {
      if (!Files.isRegularFile(file) || !Files.readString(file).equals(expected)) {
        throw new IllegalStateException("Generated reference differs: " + file);
      }
      return;
    }
    Files.createDirectories(file.getParent());
    Files.writeString(file, expected);
  }

  private static void removeObsoletePages(Path root, Map<Path, String> documents, boolean verify)
      throws IOException {
    Path reference = root.resolve("docs/reference");
    if (!Files.isDirectory(reference)) return;
    try (var files = Files.list(reference)) {
      for (Path file : files.filter(path -> path.toString().endsWith(".md")).toList()) {
        if (documents.containsKey(file)) continue;
        if (verify)
          throw new IllegalStateException("Obsolete reference page: " + file.getFileName());
        Files.delete(file);
      }
    }
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap") // One invocation owns this ordered output index.
  static Map<Path, String> render(Path root) throws IOException {
    var catalog = GridGrindProtocolCatalog.catalog();
    Map<Path, String> documents = new LinkedHashMap<>();
    List<String> index =
        new ArrayList<>(List.of(GENERATED, "# Protocol reference", "", "- [Request](request.md)"));
    documents.put(
        root.resolve("docs/reference/request.md"), page("Request", List.of(catalog.requestType())));
    for (var group : catalog.topLevelGroups())
      addGroup(root, documents, index, group.group(), group.types());
    for (var group : catalog.nestedTypes())
      addGroup(root, documents, index, group.group(), group.types());
    for (var group : catalog.plainTypes())
      addGroup(root, documents, index, group.group(), List.of(group.type()));
    index.add(
        "\nJSON Schemas: [request](schema/request.schema.json), [grant](schema/grant.schema.json), [response](schema/response.schema.json).");
    documents.put(root.resolve("docs/reference/README.md"), lines(index));
    documents.put(root.resolve("docs/LIMITATIONS.md"), limits(root));
    documents.put(root.resolve("docs/ERRORS.md"), errors());
    return documents;
  }

  private static void addGroup(
      Path root,
      Map<Path, String> documents,
      List<String> index,
      String group,
      List<TypeEntry> types)
      throws IOException {
    String name = group + ".md";
    if (documents.putIfAbsent(root.resolve("docs/reference").resolve(name), page(group, types))
        != null) {
      throw new IllegalStateException("Duplicate reference group: " + group);
    }
    index.add("- [" + group + "](" + name + ")");
  }

  private static String page(String title, List<TypeEntry> types) throws IOException {
    List<String> text = new ArrayList<>(List.of(GENERATED, "# " + title));
    for (TypeEntry type : types) {
      text.addAll(
          List.of(
              "",
              "## " + type.id(),
              "",
              type.summary(),
              "",
              "| Field | Requirement | Shape | Default |",
              "| --- | --- | --- | --- |"));
      for (var field : type.fields()) {
        String shape = GridGrindJsonMapperSupport.JSON_MAPPER.writeValueAsString(field.shape());
        String defaultValue =
            field
                .defaultFrom()
                .map(value -> "from " + value)
                .orElseGet(() -> field.defaultValue().map(Object::toString).orElse(""));
        text.add(
            "| "
                + cell(field.name())
                + " | "
                + field.requirement()
                + " | "
                + cell(shape)
                + " | "
                + cell(defaultValue)
                + " |");
      }
      if (!type.targetSelectors().isEmpty()) text.add("\nTargets: " + type.targetSelectors() + ".");
      if (!type.preconditions().isEmpty())
        text.add("\nPreconditions: " + type.preconditions() + ".");
      text.addAll(
          List.of(
              "",
              FENCE + "json",
              GridGrindJsonMapperSupport.JSON_MAPPER
                  .writerWithDefaultPrettyPrinter()
                  .writeValueAsString(type),
              FENCE));
    }
    return lines(text);
  }

  private static String limits(Path root) throws IOException {
    Map<String, List<String>> owners = enforcementOwners(root);
    List<String> text =
        new ArrayList<>(
            List.of(
                GENERATED,
                "# Limits",
                "",
                "These entries are owned by core.limits.Limit. Numeric enforcement, help and catalog publication use that owner."));
    for (Limit limit : Limit.all()) {
      text.addAll(
          List.of(
              "",
              "## " + limit.id() + " — " + limit.title(),
              "",
              "| Field | Value |",
              "| --- | --- |",
              "| Category | " + cell(limit.category()) + " |",
              "| Limit | " + cell(limit.value()) + " |",
              "| Problems | " + cell(String.join(", ", limit.problemCodes())) + " |",
              "| Message template | " + cell(limit.messageTemplate()) + " |",
              "| Applies to | " + cell(String.join(", ", limit.appliesTo())) + " |",
              "",
              limit.rationale()));
      if (owners.containsKey(limit.id()))
        text.add("\nEnforcement: " + String.join(", ", owners.get(limit.id())) + ".");
    }
    return lines(text);
  }

  @SuppressWarnings({"PMD.UseConcurrentHashMap", "PMD.AvoidInstantiatingObjectsInLoops"})
  // A local sorted relation and one owner-list per ID are required output data, not shared state.
  private static Map<String, List<String>> enforcementOwners(Path root) throws IOException {
    Map<String, List<String>> owners = new java.util.TreeMap<>();
    for (String module : List.of("core", "json", "engine")) {
      Path sources = root.resolve(module + "/src/main/java");
      try (var paths = Files.walk(sources)) {
        for (Path path :
            paths.filter(file -> file.toString().endsWith(".java")).sorted().toList()) {
          if (!Files.readString(path).contains("@dev.erst.gridgrind.core.limits.EnforcesLimit"))
            continue;
          String relative = sources.relativize(path).toString();
          String name =
              relative.substring(0, relative.length() - 5).replace(java.io.File.separatorChar, '.');
          try {
            Class<?> type =
                Class.forName(name, false, Thread.currentThread().getContextClassLoader());
            EnforcesLimit annotation = type.getAnnotation(EnforcesLimit.class);
            if (annotation == null)
              throw new IllegalStateException("Uncompiled limit annotation: " + name);
            for (String id : annotation.value()) {
              if (Limit.find(id).isEmpty())
                throw new IllegalStateException("Unknown enforced limit: " + id);
              String link =
                  "["
                      + type.getSimpleName()
                      + "](../"
                      + root.relativize(path).toString().replace(java.io.File.separatorChar, '/')
                      + ")";
              owners.computeIfAbsent(id, ignored -> new ArrayList<>()).add(link);
            }
          } catch (ClassNotFoundException exception) {
            throw new IOException("Uncompiled enforcement owner: " + name, exception);
          }
        }
      }
    }
    return owners;
  }

  private static String errors() {
    List<String> text = new ArrayList<>(List.of(GENERATED, "# Problems and recovery", ""));
    for (GridGrindProblemCode code : GridGrindProblemCode.values()) {
      text.addAll(
          List.of(
              "## " + code.name(),
              "",
              code.title(),
              "",
              "Category: " + code.category() + ". Recovery: " + code.recovery() + ".",
              "",
              code.resolution(),
              ""));
    }
    return lines(text);
  }

  private static String lines(List<String> lines) {
    return lines.stream()
            .map(String::stripTrailing)
            .collect(java.util.stream.Collectors.joining("\n"))
            .stripTrailing()
        + "\n";
  }

  private static String cell(String value) {
    return value.replace("|", "\\|").replace("\n", "<br>");
  }
}
