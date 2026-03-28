package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Complete GridGrind request for workbook source, mutations, reads, and persistence. */
public record GridGrindRequest(
    GridGrindProtocolVersion protocolVersion,
    WorkbookSource source,
    WorkbookPersistence persistence,
    List<WorkbookOperation> operations,
    List<WorkbookReadOperation> reads) {
  public GridGrindRequest {
    protocolVersion =
        protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
    Objects.requireNonNull(source, "source must not be null");
    persistence = persistence == null ? new WorkbookPersistence.None() : persistence;
    operations = copyOperations(operations);
    reads = copyReads(reads);
  }

  /** Creates a request with the current protocol version and the given parameters. */
  public GridGrindRequest(
      WorkbookSource source,
      WorkbookPersistence persistence,
      List<WorkbookOperation> operations,
      List<WorkbookReadOperation> reads) {
    this(GridGrindProtocolVersion.current(), source, persistence, operations, reads);
  }

  /** Describes where the input workbook comes from. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookSource.New.class, name = "NEW"),
    @JsonSubTypes.Type(value = WorkbookSource.ExistingFile.class, name = "EXISTING")
  })
  public sealed interface WorkbookSource {
    /** Creates a brand-new empty workbook in memory. */
    record New() implements WorkbookSource {}

    /** Opens an existing workbook file from disk. */
    record ExistingFile(String path) implements WorkbookSource {
      public ExistingFile {
        requireXlsxWorkbookPath(path);
      }
    }
  }

  /** Describes whether and where the resulting workbook should be saved. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookPersistence.None.class, name = "NONE"),
    @JsonSubTypes.Type(value = WorkbookPersistence.OverwriteSource.class, name = "OVERWRITE"),
    @JsonSubTypes.Type(value = WorkbookPersistence.SaveAs.class, name = "SAVE_AS")
  })
  public sealed interface WorkbookPersistence {
    /** Leaves the workbook in memory only and does not persist it. */
    record None() implements WorkbookPersistence {}

    /** Saves the workbook back to the exact path it was opened from. */
    record OverwriteSource() implements WorkbookPersistence {}

    /** Saves the workbook to a new `.xlsx` path. */
    record SaveAs(String path) implements WorkbookPersistence {
      public SaveAs {
        requireXlsxWorkbookPath(path);
      }
    }
  }

  private static List<WorkbookOperation> copyOperations(List<WorkbookOperation> operations) {
    if (operations == null) {
      return List.of();
    }
    List<WorkbookOperation> copy = List.copyOf(operations);
    for (WorkbookOperation operation : copy) {
      Objects.requireNonNull(operation, "operations must not contain nulls");
    }
    return copy;
  }

  private static List<WorkbookReadOperation> copyReads(List<WorkbookReadOperation> reads) {
    if (reads == null) {
      return List.of();
    }
    List<WorkbookReadOperation> copy = List.copyOf(reads);
    for (WorkbookReadOperation read : copy) {
      Objects.requireNonNull(read, "reads must not contain nulls");
    }
    return copy;
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  private static void requireXlsxWorkbookPath(String path) {
    requireNonBlank(path, "path");
    if (!path.toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
      throw new IllegalArgumentException(
          "path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: " + path);
    }
  }
}
