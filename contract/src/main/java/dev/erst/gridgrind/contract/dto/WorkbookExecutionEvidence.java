package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.catalog.OperationEffect;
import dev.erst.gridgrind.contract.catalog.OperationEffectFootprint;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Typed evidence established while executing one workbook plan. */
public record WorkbookExecutionEvidence(
    Admission admission,
    Structural structural,
    Computational computational,
    TaskSpecific taskSpecific,
    Presentational presentational,
    List<Preservation> preservation) {
  public WorkbookExecutionEvidence {
    Objects.requireNonNull(admission, "admission must not be null");
    Objects.requireNonNull(structural, "structural must not be null");
    Objects.requireNonNull(computational, "computational must not be null");
    Objects.requireNonNull(taskSpecific, "taskSpecific must not be null");
    Objects.requireNonNull(presentational, "presentational must not be null");
    preservation =
        List.copyOf(Objects.requireNonNull(preservation, "preservation must not be null"));
  }

  /** Effects and immutable input identities admitted for the executed plan. */
  public record Admission(List<Operation> operations, List<MaterializedInput> materializedInputs) {
    public Admission {
      operations = copyValues(operations, "operations");
      materializedInputs = copyValues(materializedInputs, "materializedInputs");
    }

    /** Returns explicit evidence for a request that did not reach materialized input binding. */
    public static Admission empty() {
      return new Admission(List.of(), List.of());
    }
  }

  /** One admitted concrete operation's canonical effects and footprint. */
  public record Operation(
      String operationId, List<OperationEffect> effects, OperationEffectFootprint footprint) {
    public Operation {
      operationId = requireNonBlank(operationId, "operationId");
      effects = copyValues(effects, "effects");
      if (effects.isEmpty()) {
        throw new IllegalArgumentException("effects must not be empty");
      }
      Objects.requireNonNull(footprint, "footprint must not be null");
    }
  }

  /** One immutable input copy identified by its bound path, byte count, and SHA-256 digest. */
  public record MaterializedInput(String role, String path, long byteSize, String sha256) {
    public MaterializedInput {
      role = requireNonBlank(role, "role");
      path = requireNonBlank(path, "path");
      if (byteSize < 0) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      sha256 = requireSha256(sha256);
    }
  }

  /** Structural claim for an artifact that was or was not staged and reopened. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Structural.NotEstablished.class, name = "NOT_ESTABLISHED"),
    @JsonSubTypes.Type(value = Structural.Established.class, name = "ESTABLISHED")
  })
  public sealed interface Structural permits Structural.NotEstablished, Structural.Established {
    record NotEstablished() implements Structural {}

    record Established() implements Structural {}
  }

  /** Calculation claim with an explicit unestablished state. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Computational.NotRequested.class, name = "NOT_REQUESTED"),
    @JsonSubTypes.Type(value = Computational.Established.class, name = "ESTABLISHED"),
    @JsonSubTypes.Type(value = Computational.Unestablished.class, name = "UNESTABLISHED")
  })
  public sealed interface Computational
      permits Computational.NotRequested, Computational.Established, Computational.Unestablished {
    record NotRequested() implements Computational {}

    record Established() implements Computational {}

    record Unestablished() implements Computational {}
  }

  /** Plan-authored and host-required assertion evidence with its authoritative origin. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = TaskSpecific.NotRequested.class, name = "NOT_REQUESTED"),
    @JsonSubTypes.Type(value = TaskSpecific.Established.class, name = "ESTABLISHED"),
    @JsonSubTypes.Type(value = TaskSpecific.Failed.class, name = "FAILED")
  })
  public sealed interface TaskSpecific
      permits TaskSpecific.NotRequested, TaskSpecific.Established, TaskSpecific.Failed {
    record NotRequested() implements TaskSpecific {}

    record Established(List<Assertion> assertions) implements TaskSpecific {
      public Established {
        assertions = copyAssertions(assertions, "assertions");
        if (assertions.isEmpty()) {
          throw new IllegalArgumentException("assertions must not be empty");
        }
        if (assertions.stream()
            .anyMatch(assertion -> assertion.result() instanceof AssertionResult.Failed)) {
          throw new IllegalArgumentException(
              "established task-specific evidence cannot contain failures");
        }
      }
    }

    record Failed(List<Assertion> assertions) implements TaskSpecific {
      public Failed {
        assertions = copyAssertions(assertions, "assertions");
        if (assertions.stream()
            .noneMatch(assertion -> assertion.result() instanceof AssertionResult.Failed)) {
          throw new IllegalArgumentException(
              "failed task-specific evidence must contain a failure");
        }
      }
    }

    /** One assertion outcome and whether the plan or trusted host required it. */
    public record Assertion(Origin origin, AssertionResult result) {
      public Assertion {
        Objects.requireNonNull(origin, "origin must not be null");
        Objects.requireNonNull(result, "result must not be null");
      }
    }

    /** The authority that required one assertion. */
    public enum Origin {
      PLAN_AUTHORED,
      HOST_REQUIRED
    }
  }

  /** Presentation is never inferred from OOXML structure. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Presentational.NotAssessed.class, name = "NOT_ASSESSED")
  })
  public sealed interface Presentational permits Presentational.NotAssessed {
    record NotAssessed() implements Presentational {}
  }

  /** Preservation claim including semantic and byte-distinct states. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Preservation.NotRequested.class, name = "NOT_REQUESTED"),
    @JsonSubTypes.Type(value = Preservation.NotApplicable.class, name = "NOT_APPLICABLE"),
    @JsonSubTypes.Type(
        value = Preservation.SemanticEstablished.class,
        name = "SEMANTIC_ESTABLISHED"),
    @JsonSubTypes.Type(value = Preservation.ByteEstablished.class, name = "BYTE_ESTABLISHED"),
    @JsonSubTypes.Type(value = Preservation.Failed.class, name = "FAILED"),
    @JsonSubTypes.Type(value = Preservation.Unestablished.class, name = "UNESTABLISHED")
  })
  public sealed interface Preservation
      permits Preservation.NotRequested,
          Preservation.NotApplicable,
          Preservation.SemanticEstablished,
          Preservation.ByteEstablished,
          Preservation.Failed,
          Preservation.Unestablished {
    record NotRequested() implements Preservation {}

    record NotApplicable() implements Preservation {}

    record SemanticEstablished(List<String> inspectionStepIds) implements Preservation {
      public SemanticEstablished {
        inspectionStepIds = copyIdentifiers(inspectionStepIds, "inspectionStepIds");
      }
    }

    record ByteEstablished(List<String> partNames) implements Preservation {
      public ByteEstablished {
        partNames = copyIdentifiers(partNames, "partNames");
      }
    }

    record Failed(Kind kind, List<String> identifiers) implements Preservation {
      public Failed {
        Objects.requireNonNull(kind, "kind must not be null");
        identifiers = copyIdentifiers(identifiers, "identifiers");
      }
    }

    record Unestablished() implements Preservation {}

    /** The preservation evidence family whose required comparison failed. */
    public enum Kind {
      SEMANTIC,
      BYTE
    }
  }

  /** Returns minimum factual evidence for the existing execution path. */
  public static WorkbookExecutionEvidence minimum() {
    return new WorkbookExecutionEvidence(
        Admission.empty(),
        new Structural.NotEstablished(),
        new Computational.NotRequested(),
        new TaskSpecific.NotRequested(),
        new Presentational.NotAssessed(),
        List.of(new Preservation.NotRequested()));
  }

  private static List<TaskSpecific.Assertion> copyAssertions(
      List<TaskSpecific.Assertion> assertions, String fieldName) {
    Objects.requireNonNull(assertions, fieldName + " must not be null");
    List<TaskSpecific.Assertion> copied = new ArrayList<>(assertions.size());
    for (TaskSpecific.Assertion assertion : assertions) {
      copied.add(Objects.requireNonNull(assertion, fieldName + " must not contain null values"));
    }
    return List.copyOf(copied);
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copied = new ArrayList<>(values.size());
    for (T value : values) {
      copied.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copied);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static String requireSha256(String value) {
    String digest = requireNonBlank(value, "sha256");
    if (!digest.matches("[0-9a-f]{64}")) {
      throw new IllegalArgumentException("sha256 must be a lowercase SHA-256 hexadecimal digest");
    }
    return digest;
  }

  private static List<String> copyIdentifiers(List<String> values, String fieldName) {
    List<String> copied = copyValues(values, fieldName);
    if (copied.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    for (String value : copied) {
      requireNonBlank(value, fieldName + " element");
    }
    return copied;
  }
}
