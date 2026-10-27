package dev.erst.gridgrind.json.schema;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.EnumMap;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import tools.jackson.databind.cfg.JsonNodeFeature;
import tools.jackson.databind.json.JsonMapper;

/** Publishes deterministic Draft 2020-12 schemas from the canonical document models. */
public final class JsonSchemas {
  private static final JsonMapper JSON =
      JsonMapper.builder().enable(JsonNodeFeature.WRITE_PROPERTIES_SORTED).build();
  private static final Map<Document, String> DOCUMENTS = documents();

  private JsonSchemas() {}

  /** The three externally exchanged document contracts. */
  public enum Document {
    REQUEST(WorkbookPlan.class),
    GRANT(ExecutionGrant.class),
    RESPONSE(WorkbookResult.class);

    private final Class<?> rootType;

    Document(Class<?> rootType) {
      this.rootType = rootType;
    }

    /** Returns the canonical model that owns this document. */
    public Class<?> rootType() {
      return rootType;
    }

    /** Returns the generated schema artifact name. */
    public String fileName() {
      return name().toLowerCase(Locale.ROOT) + ".schema.json";
    }
  }

  /** Returns an immutable JSON schema document; callers cannot mutate shared schema state. */
  public static String document(Document document) {
    return Objects.requireNonNull(DOCUMENTS.get(document));
  }

  @SuppressWarnings(
      "PMD.UseConcurrentHashMap") // Constructed locally, then published as an immutable map.
  private static Map<Document, String> documents() {
    Map<Document, String> documents = new EnumMap<>(Document.class);
    for (Document document : Document.values()) {
      var schema = SchemaGenerator.generate(document.rootType(), document == Document.RESPONSE);
      documents.put(
          document, JSON.writerWithDefaultPrettyPrinter().writeValueAsString(schema) + "\n");
    }
    return Map.copyOf(documents);
  }

  /** Writes the generated reference artifacts for the build's generateSchemas task. */
  @SuppressWarnings(
      "PMD.SystemPrintln") // The build-tool entrypoint reports the generated reference paths.
  public static void main(String[] args) throws IOException {
    if (args.length != 1) {
      throw new IllegalArgumentException("Usage: JsonSchemas <output-directory>");
    }
    Path directory = Path.of(args[0]);
    Files.createDirectories(directory);
    for (Document document : Document.values()) {
      Path output = directory.resolve(document.fileName());
      Files.writeString(output, document(document), StandardCharsets.UTF_8);
      System.out.println("Generated " + output);
    }
  }
}
