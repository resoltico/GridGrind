package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import java.util.Objects;

/** The authored-only image boundary; observation and raw part copying never call this owner. */
public final class AuthoredImageAdmission {
  private AuthoredImageAdmission() {}

  /** Inspects materialized immutable content against its request's shared counters. */
  public static void inspect(
      ExcelPictureFormat format, BinaryContent content, MediaAdmissionBudget budget) {
    budget.requirePayloadSize(content.size());
    check(format, content.bytes(), budget);
  }

  /** Guards a direct native receiving operation before any destructive work. */
  public static AdmittedImage admitNative(ExcelPictureFormat format, BinaryContent content) {
    Objects.requireNonNull(format);
    Objects.requireNonNull(content);
    var budget = new MediaAdmissionBudget();
    budget.reserveOccurrence(content.size());
    byte[] owned = content.bytes();
    check(format, owned, budget);
    return new AdmittedImage(format, content, owned);
  }

  private static void check(ExcelPictureFormat format, byte[] bytes, MediaAdmissionBudget budget) {
    try {
      budget.inspect(bytes.length);
      ImageStructures.check(new MediaInput(bytes, 0, bytes.length, format, budget, 0));
    } catch (InvalidRequestException expected) {
      if (expected.requestProblem() instanceof FieldValidationProblem field
          && field.rule() instanceof FieldValidationMediaRule) throw expected;
      throw new UnexpectedNativeContentException(expected);
    } catch (UnexpectedNativeContentException classified) {
      throw classified;
    } catch (RuntimeException unexpected) {
      throw new UnexpectedNativeContentException(unexpected);
    }
  }

  /** Keeps unclassified failures at the authored native mutation role unexpected. */
  public static void mutate(Runnable mutation) {
    try {
      mutation.run();
    } catch (UnexpectedNativeContentException classified) {
      throw classified;
    } catch (RuntimeException unexpected) {
      throw new UnexpectedNativeContentException(unexpected);
    }
  }

  /** Privately owned structural admission evidence, never a claim of universal rendering. */
  public static final class AdmittedImage {
    private final ExcelPictureFormat format;
    private final BinaryContent source;
    private final byte[] bytes;

    private AdmittedImage(ExcelPictureFormat format, BinaryContent source, byte[] bytes) {
      this.format = format;
      this.source = source;
      this.bytes = bytes;
    }

    /** Returns an independent native copy only for the exact immutable source that was admitted. */
    @SuppressWarnings({"ReferenceEquality", "PMD.CompareObjectsWithEquals"})
    // This checks ownership of the immutable source instance, not equal-payload identity caching.
    public byte[] nativeBytes(ExcelPictureFormat expectedFormat, BinaryContent expectedSource) {
      if (format != expectedFormat || source != expectedSource)
        throw new UnexpectedNativeContentException(
            new IllegalStateException("Admitted image ownership mismatch"));
      return bytes.clone();
    }
  }
}
