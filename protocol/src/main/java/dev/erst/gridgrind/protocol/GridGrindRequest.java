package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;

/** Complete agent request for workbook source, mutations, persistence, and analysis. */
public record GridGrindRequest(
    GridGrindProtocolVersion protocolVersion,
    WorkbookSource source,
    WorkbookPersistence persistence,
    List<WorkbookOperation> operations,
    WorkbookAnalysisRequest analysis) {
  public GridGrindRequest {
    protocolVersion = protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
    Objects.requireNonNull(source, "source must not be null");
    persistence = persistence == null ? new WorkbookPersistence.None() : persistence;
    operations = copyOperations(operations);
    analysis = analysis == null ? new WorkbookAnalysisRequest(List.of()) : analysis;
  }

  public GridGrindRequest(
      WorkbookSource source,
      WorkbookPersistence persistence,
      List<WorkbookOperation> operations,
      WorkbookAnalysisRequest analysis) {
    this(GridGrindProtocolVersion.current(), source, persistence, operations, analysis);
  }

  /** Describes where the input workbook comes from. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "mode")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookSource.New.class, name = "NEW"),
    @JsonSubTypes.Type(value = WorkbookSource.ExistingFile.class, name = "EXISTING_FILE")
  })
  public sealed interface WorkbookSource {
    record New() implements WorkbookSource {}

    record ExistingFile(String path) implements WorkbookSource {
      public ExistingFile {
        requireNonBlank(path, "path");
      }
    }
  }

  /** Describes whether and where the resulting workbook should be saved. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "mode")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookPersistence.None.class, name = "NONE"),
    @JsonSubTypes.Type(
        value = WorkbookPersistence.OverwriteSource.class,
        name = "OVERWRITE_SOURCE"),
    @JsonSubTypes.Type(value = WorkbookPersistence.SaveAs.class, name = "SAVE_AS")
  })
  public sealed interface WorkbookPersistence {
    record None() implements WorkbookPersistence {}

    record OverwriteSource() implements WorkbookPersistence {}

    record SaveAs(String path) implements WorkbookPersistence {
      public SaveAs {
        requireNonBlank(path, "path");
      }
    }
  }

  /** Declares which workbook content should be inspected after mutations are applied. */
  public record WorkbookAnalysisRequest(List<SheetInspectionRequest> sheets) {
    public WorkbookAnalysisRequest {
      sheets = copySheets(sheets);
    }
  }

  /** Declares how one sheet should be inspected after mutations finish. */
  public record SheetInspectionRequest(
      String sheetName, List<String> cells, Integer previewRowCount, Integer previewColumnCount) {
    public SheetInspectionRequest {
      requireNonBlank(sheetName, "sheetName");
      cells = cells == null ? List.of() : List.copyOf(cells);
      for (String cell : cells) {
        requireNonBlank(cell, "cell");
      }
      if ((previewRowCount == null) != (previewColumnCount == null)) {
        throw new IllegalArgumentException(
            "previewRowCount and previewColumnCount must both be set or both be null");
      }
      if (previewRowCount != null && previewRowCount <= 0) {
        throw new IllegalArgumentException("previewRowCount must be greater than 0");
      }
      if (previewColumnCount != null && previewColumnCount <= 0) {
        throw new IllegalArgumentException("previewColumnCount must be greater than 0");
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

  private static List<SheetInspectionRequest> copySheets(List<SheetInspectionRequest> sheets) {
    if (sheets == null) {
      return List.of();
    }
    List<SheetInspectionRequest> copy = List.copyOf(sheets);
    for (SheetInspectionRequest sheet : copy) {
      Objects.requireNonNull(sheet, "sheets must not contain nulls");
    }
    return copy;
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }
}
