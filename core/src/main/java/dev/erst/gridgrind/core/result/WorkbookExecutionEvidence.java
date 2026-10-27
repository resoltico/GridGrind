package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.semantics.OperationEffect;
import dev.erst.gridgrind.core.semantics.OperationEffectFootprint;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
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
      operationId = TextValidation.requireNonBlank(operationId, "operationId");
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
      role = TextValidation.requireNonBlank(role, "role");
      path = TextValidation.requireNonBlank(path, "path");
      if (byteSize < 0) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      sha256 = requireSha256(sha256);
    }
  }

  /** Structural claim for an artifact that was or was not staged and reopened. */
  @Union(discriminator = "status")
  public sealed interface Structural permits Structural.NotEstablished, Structural.Established {
    @Variant(id = "NOT_ESTABLISHED", summary = "Structural verification has not been established.")
    record NotEstablished() implements Structural {}

    @Variant(
        id = "ESTABLISHED",
        summary = "The staged artifact was reopened and structural verification succeeded.")
    record Established() implements Structural {}
  }

  /** Calculation claim with an explicit unestablished state. */
  @Union(discriminator = "status")
  public sealed interface Computational
      permits Computational.NotRequested, Computational.Established, Computational.Unestablished {
    @Variant(id = "NOT_REQUESTED", summary = "No server-side calculation was requested.")
    record NotRequested() implements Computational {}

    @Variant(id = "ESTABLISHED", summary = "The requested calculation capability was established.")
    record Established() implements Computational {}

    @Variant(
        id = "UNESTABLISHED",
        summary = "The requested calculation capability was not established.")
    record Unestablished() implements Computational {}
  }

  /** Plan-authored and host-required assertion evidence with its authoritative origin. */
  @Union(discriminator = "status")
  public sealed interface TaskSpecific
      permits TaskSpecific.NotRequested, TaskSpecific.Established, TaskSpecific.Failed {
    @Variant(id = "NOT_REQUESTED", summary = "No task-specific assertion evidence was requested.")
    record NotRequested() implements TaskSpecific {}

    @Variant(id = "ESTABLISHED", summary = "Every recorded task-specific assertion passed.")
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

    @Variant(id = "FAILED", summary = "A recorded task-specific assertion failed.")
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
  @Union(discriminator = "status")
  public sealed interface Presentational permits Presentational.NotAssessed {
    @Variant(id = "NOT_ASSESSED", summary = "Presentational correctness was not assessed.")
    record NotAssessed() implements Presentational {}
  }

  /** Preservation claim including semantic and byte-distinct states. */
  @Union(discriminator = "status")
  public sealed interface Preservation
      permits Preservation.NotRequested,
          Preservation.NotApplicable,
          Preservation.SemanticEstablished,
          Preservation.ByteEstablished,
          Preservation.Failed,
          Preservation.Unestablished {
    @Variant(id = "NOT_REQUESTED", summary = "No preservation verification was requested.")
    record NotRequested() implements Preservation {}

    @Variant(
        id = "NOT_APPLICABLE",
        summary = "Preservation verification is not applicable to this execution.")
    record NotApplicable() implements Preservation {}

    @Variant(
        id = "SEMANTIC_ESTABLISHED",
        summary = "Semantic preservation was established by the reported inspections.")
    record SemanticEstablished(List<dev.erst.gridgrind.core.vocabulary.StepId> inspectionStepIds)
        implements Preservation {
      public SemanticEstablished {
        inspectionStepIds = copyRequiredValues(inspectionStepIds, "inspectionStepIds");
      }
    }

    @Variant(
        id = "BYTE_ESTABLISHED",
        summary = "Byte preservation was established for the reported package parts.")
    record ByteEstablished(List<String> partNames) implements Preservation {
      public ByteEstablished {
        partNames = copyIdentifiers(partNames, "partNames");
      }
    }

    @Variant(id = "FAILED", summary = "Requested preservation verification failed.")
    record Failed(Kind kind, List<String> identifiers) implements Preservation {
      public Failed {
        Objects.requireNonNull(kind, "kind must not be null");
        identifiers = copyIdentifiers(identifiers, "identifiers");
      }
    }

    @Variant(
        id = "UNESTABLISHED",
        summary = "Requested preservation verification was not established.")
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

  private static String requireSha256(String value) {
    String digest = TextValidation.requireNonBlank(value, "sha256");
    if (!digest.matches("[0-9a-f]{64}")) {
      throw new IllegalArgumentException("sha256 must be a lowercase SHA-256 hexadecimal digest");
    }
    return digest;
  }

  private static <T> List<T> copyRequiredValues(List<T> values, String fieldName) {
    List<T> copied = copyValues(values, fieldName);
    if (copied.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return copied;
  }

  private static List<String> copyIdentifiers(List<String> values, String fieldName) {
    List<String> copied = copyRequiredValues(values, fieldName);
    for (String value : copied) {
      TextValidation.requireNonBlank(value, fieldName + " element");
    }
    return copied;
  }
}
