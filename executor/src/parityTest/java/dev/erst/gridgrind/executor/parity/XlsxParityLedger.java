package dev.erst.gridgrind.executor.parity;

import java.io.IOException;
import java.io.InputStream;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import tools.jackson.databind.DeserializationFeature;
import tools.jackson.databind.json.JsonMapper;

/** Loads the tracked Apache POI XSSF parity ledger and workbook-corpus manifest. */
public final class XlsxParityLedger {
  private static final JsonMapper JSON_MAPPER =
      JsonMapper.builder().disable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES).build();
  private static final String LEDGER_RESOURCE =
      "/dev/erst/gridgrind/executor/parity/xssf-parity-ledger.json";
  private static final String CORPUS_RESOURCE =
      "/dev/erst/gridgrind/executor/parity/xssf-parity-corpus.json";

  private XlsxParityLedger() {}

  /** Returns the canonical tracked parity ledger. */
  public static Ledger loadLedger() throws IOException {
    return load(LEDGER_RESOURCE, Ledger.class);
  }

  /** Returns the canonical tracked workbook-corpus manifest. */
  public static CorpusManifest loadCorpus() throws IOException {
    return load(CORPUS_RESOURCE, CorpusManifest.class);
  }

  private static <T> T load(String resourcePath, Class<T> type) throws IOException {
    try (InputStream inputStream = XlsxParityLedger.class.getResourceAsStream(resourcePath)) {
      if (inputStream == null) {
        throw new IOException("Missing parity resource: " + resourcePath);
      }
      return JSON_MAPPER.readValue(inputStream, type);
    }
  }

  /** Machine-readable parity ledger loaded from the committed JSON resource. */
  public record Ledger(String version, String poiVersion, List<Row> rows) {
    public Ledger {
      version = requireNonBlank(version, "version");
      poiVersion = requireNonBlank(poiVersion, "poiVersion");
      rows = List.copyOf(Objects.requireNonNull(rows, "rows must not be null"));
      if (rows.isEmpty()) {
        throw new IllegalArgumentException("rows must not be empty");
      }
      requireUnique(
          rows.stream().map(Row::sequence).toList(), "row sequence values must be unique");
      requireUnique(rows.stream().map(Row::id).toList(), "row ids must be unique");
      List<Integer> orderedSequences =
          rows.stream().map(Row::sequence).sorted(Comparator.naturalOrder()).toList();
      for (int index = 0; index < orderedSequences.size(); index++) {
        int expectedSequence = index + 1;
        if (orderedSequences.get(index) != expectedSequence) {
          throw new IllegalArgumentException(
              "row sequences must be contiguous starting at 1; expected "
                  + expectedSequence
                  + " but found "
                  + orderedSequences.get(index));
        }
      }
    }

    /** Returns the row with the given stable identifier. */
    public Row row(String id) {
      Objects.requireNonNull(id, "id must not be null");
      return rows.stream()
          .filter(row -> row.id().equals(id))
          .findFirst()
          .orElseThrow(() -> new IllegalArgumentException("Unknown parity row id: " + id));
    }

    /** Returns the ledger rows keyed by stable identifier. */
    public Map<String, Row> rowsById() {
      Map<String, Row> rowsById =
          rows.stream()
              .collect(
                  Collectors.toMap(Row::id, row -> row, (left, right) -> left, LinkedHashMap::new));
      return Map.copyOf(rowsById);
    }
  }

  /** One semantic parity row owned by exactly one blocking phase. */
  public record Row(
      int sequence,
      String id,
      String capability,
      String subfamily,
      int phase,
      PoiSupport poiSupport,
      GridGrindStatus gridGrindStatus,
      ExpectedOutcome expectedOutcome,
      String probeId,
      List<String> scenarioIds,
      List<String> primarySources,
      String notes) {
    public Row {
      if (sequence <= 0) {
        throw new IllegalArgumentException("sequence must be greater than 0");
      }
      id = requireNonBlank(id, "id");
      capability = requireNonBlank(capability, "capability");
      subfamily = requireNonBlank(subfamily, "subfamily");
      if (phase < 1 || phase > 10) {
        throw new IllegalArgumentException("phase must be between 1 and 10 inclusive: " + phase);
      }
      Objects.requireNonNull(poiSupport, "poiSupport must not be null");
      Objects.requireNonNull(gridGrindStatus, "gridGrindStatus must not be null");
      Objects.requireNonNull(expectedOutcome, "expectedOutcome must not be null");
      probeId = requireNonBlank(probeId, "probeId");
      scenarioIds = copyDistinctNonBlankStrings(scenarioIds, "scenarioIds");
      primarySources = copyDistinctNonBlankStrings(primarySources, "primarySources");
      notes = notes == null ? "" : notes;
    }
  }

  /** Machine-readable workbook-corpus manifest loaded from the committed JSON resource. */
  public record CorpusManifest(String version, String poiVersion, List<Scenario> scenarios) {
    public CorpusManifest {
      version = requireNonBlank(version, "version");
      poiVersion = requireNonBlank(poiVersion, "poiVersion");
      scenarios = List.copyOf(Objects.requireNonNull(scenarios, "scenarios must not be null"));
      if (scenarios.isEmpty()) {
        throw new IllegalArgumentException("scenarios must not be empty");
      }
      requireUnique(scenarios.stream().map(Scenario::id).toList(), "scenario ids must be unique");
    }

    /** Returns the committed scenario descriptor with the given stable identifier. */
    public Scenario scenario(String id) {
      Objects.requireNonNull(id, "id must not be null");
      return scenarios.stream()
          .filter(scenario -> scenario.id().equals(id))
          .findFirst()
          .orElseThrow(() -> new IllegalArgumentException("Unknown parity scenario id: " + id));
    }

    /** Returns the scenario descriptors keyed by stable identifier. */
    public Map<String, Scenario> scenariosById() {
      Map<String, Scenario> scenariosById =
          scenarios.stream()
              .collect(
                  Collectors.toMap(
                      Scenario::id,
                      scenario -> scenario,
                      (left, right) -> left,
                      LinkedHashMap::new));
      return Map.copyOf(scenariosById);
    }
  }

  /** One deterministic corpus workbook descriptor. */
  public record Scenario(String id, ScenarioKind kind, String description) {
    public Scenario {
      id = requireNonBlank(id, "id");
      Objects.requireNonNull(kind, "kind must not be null");
      description = requireNonBlank(description, "description");
    }
  }

  /** Apache POI support classification for one parity row. */
  public enum PoiSupport {
    DOCUMENTED_SUPPORTED,
    DOCUMENTED_LIMITED_SUPPORT,
    POI_LIMITATION,
  }

  /** Current measured GridGrind support status for one parity row. */
  public enum GridGrindStatus {
    COMPLETE,
    PARTIAL,
    ABSENT,
  }

  /** Expected executable probe outcome for the current repository state. */
  public enum ExpectedOutcome {
    PASS,
    FAIL,
  }

  /** Committed corpus scenario kind. */
  public enum ScenarioKind {
    GRIDGRIND_REQUEST,
    POI_DIRECT,
    POI_SIGNED,
    POI_ENCRYPTED,
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static List<String> copyDistinctNonBlankStrings(List<String> values, String fieldName) {
    List<String> copy =
        List.copyOf(Objects.requireNonNull(values, fieldName + " must not be null"));
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    requireUnique(copy, fieldName + " must not contain duplicates");
    for (String value : copy) {
      requireNonBlank(value, fieldName);
    }
    return copy;
  }

  private static <T> void requireUnique(List<T> values, String message) {
    Set<T> unique = Set.copyOf(values);
    if (unique.size() != values.size()) {
      throw new IllegalArgumentException(message);
    }
  }
}
