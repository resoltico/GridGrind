package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;

/** Workbook-surface report families returned by inspections and successful execution. */
public interface GridGrindWorkbookSurfaceReports {
  /** High-level workbook facts returned on success. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookSummary.Empty.class, name = "EMPTY"),
    @JsonSubTypes.Type(value = WorkbookSummary.WithSheets.class, name = "WITH_SHEETS")
  })
  sealed interface WorkbookSummary permits WorkbookSummary.Empty, WorkbookSummary.WithSheets {
    /** Total sheet count after all mutations complete. */
    int sheetCount();

    /** Ordered workbook sheet names. */
    List<String> sheetNames();

    /** Count of exposed named ranges after all mutations complete. */
    int namedRangeCount();

    /** Whether the workbook is marked to recalculate formulas on open. */
    boolean forceFormulaRecalculationOnOpen();

    /** Workbook summary for a zero-sheet workbook. */
    record Empty(
        int sheetCount,
        List<String> sheetNames,
        int namedRangeCount,
        boolean forceFormulaRecalculationOnOpen)
        implements WorkbookSummary {
      public Empty {
        sheetNames =
            GridGrindResponseSupport.validateCommonWorkbookSummaryFields(
                sheetCount, sheetNames, namedRangeCount);
        if (sheetCount != 0) {
          throw new IllegalArgumentException("sheetCount must be 0 for an empty workbook");
        }
      }
    }

    /** Workbook summary for a workbook that contains one or more sheets. */
    record WithSheets(
        int sheetCount,
        List<String> sheetNames,
        String activeSheetName,
        List<String> selectedSheetNames,
        int namedRangeCount,
        boolean forceFormulaRecalculationOnOpen)
        implements WorkbookSummary {
      public WithSheets {
        sheetNames =
            GridGrindResponseSupport.validateCommonWorkbookSummaryFields(
                sheetCount, sheetNames, namedRangeCount);
        Objects.requireNonNull(activeSheetName, "activeSheetName must not be null");
        if (activeSheetName.isBlank()) {
          throw new IllegalArgumentException("activeSheetName must not be blank");
        }
        selectedSheetNames =
            GridGrindResponseSupport.copyDistinctStrings(selectedSheetNames, "selectedSheetNames");
        if (sheetCount == 0) {
          throw new IllegalArgumentException("sheetCount must be greater than 0");
        }
        if (!sheetNames.contains(activeSheetName)) {
          throw new IllegalArgumentException("activeSheetName must be present in sheetNames");
        }
        if (selectedSheetNames.isEmpty()) {
          throw new IllegalArgumentException("selectedSheetNames must not be empty");
        }
        for (String selectedSheetName : selectedSheetNames) {
          if (!sheetNames.contains(selectedSheetName)) {
            throw new IllegalArgumentException(
                "selectedSheetNames must only contain values present in sheetNames");
          }
        }
      }
    }
  }

  /** Structured workbook-level analysis report for one defined name. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = NamedRangeReport.RangeReport.class, name = "RANGE"),
    @JsonSubTypes.Type(value = NamedRangeReport.FormulaReport.class, name = "FORMULA")
  })
  sealed interface NamedRangeReport {
    /** Defined-name identifier. */
    String name();

    /** Workbook or sheet scope of the defined name. */
    NamedRangeScope scope();

    /** Exact formula text stored in the workbook for this defined name. */
    String refersToFormula();

    /** Named range that resolves cleanly to a sheet-qualified cell or rectangular range target. */
    record RangeReport(
        String name, NamedRangeScope scope, String refersToFormula, NamedRangeTarget target)
        implements NamedRangeReport {
      public RangeReport {
        Objects.requireNonNull(name, "name must not be null");
        if (name.isBlank()) {
          throw new IllegalArgumentException("name must not be blank");
        }
        Objects.requireNonNull(scope, "scope must not be null");
        Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
        if (refersToFormula.isBlank()) {
          throw new IllegalArgumentException("refersToFormula must not be blank");
        }
        Objects.requireNonNull(target, "target must not be null");
      }
    }

    /** Defined name whose formula cannot be normalized to a typed range target. */
    record FormulaReport(String name, NamedRangeScope scope, String refersToFormula)
        implements NamedRangeReport {
      public FormulaReport {
        Objects.requireNonNull(name, "name must not be null");
        if (name.isBlank()) {
          throw new IllegalArgumentException("name must not be blank");
        }
        Objects.requireNonNull(scope, "scope must not be null");
        Objects.requireNonNull(refersToFormula, "refersToFormula must not be null");
        if (refersToFormula.isBlank()) {
          throw new IllegalArgumentException("refersToFormula must not be blank");
        }
      }
    }
  }

  /** Structural summary facts for one sheet. */
  record SheetSummaryReport(
      String sheetName,
      ExcelSheetVisibility visibility,
      SheetProtectionReport protection,
      int physicalRowCount,
      int lastRowIndex,
      int lastColumnIndex) {
    public SheetSummaryReport {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      Objects.requireNonNull(visibility, "visibility must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Captures whether a sheet is protected and, if so, with which supported lock flags. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = SheetProtectionReport.Unprotected.class, name = "UNPROTECTED"),
    @JsonSubTypes.Type(value = SheetProtectionReport.Protected.class, name = "PROTECTED")
  })
  sealed interface SheetProtectionReport
      permits SheetProtectionReport.Unprotected, SheetProtectionReport.Protected {

    /** Sheet protection is disabled. */
    record Unprotected() implements SheetProtectionReport {}

    /** Sheet protection is enabled with the reported supported lock flags. */
    record Protected(SheetProtectionSettings settings) implements SheetProtectionReport {
      public Protected {
        Objects.requireNonNull(settings, "settings must not be null");
      }
    }
  }

  /** Factual comment metadata returned for analyzed cells that carry a comment. */
  record CommentReport(
      String text,
      String author,
      boolean visible,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<List<RichTextRunReport>> runs,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentAnchorReport> anchor) {
    /** Creates a plain comment report without rich runs or anchor metadata. */
    public CommentReport(String text, String author, boolean visible) {
      this(text, author, visible, Optional.empty(), Optional.empty());
    }

    public CommentReport {
      Objects.requireNonNull(text, "text must not be null");
      Objects.requireNonNull(author, "author must not be null");
      runs = GridGrindResponseSupport.copyOptionalValues(runs, "runs");
      anchor = Objects.requireNonNullElseGet(anchor, Optional::empty);
      if (runs.isPresent()) {
        List<RichTextRunReport> copiedRuns = runs.orElseThrow();
        if (!text.equals(
            copiedRuns.stream().map(RichTextRunReport::text).collect(Collectors.joining()))) {
          throw new IllegalArgumentException("comment runs must concatenate to the plain text");
        }
      }
    }
  }

  /** Effective style facts returned with every analyzed cell. */
  record CellStyleReport(
      String numberFormat,
      CellAlignmentReport alignment,
      CellFontReport font,
      CellFillReport fill,
      CellBorderReport border,
      CellProtectionReport protection) {
    public CellStyleReport {
      Objects.requireNonNull(numberFormat, "numberFormat must not be null");
      Objects.requireNonNull(alignment, "alignment must not be null");
      Objects.requireNonNull(font, "font must not be null");
      Objects.requireNonNull(fill, "fill must not be null");
      Objects.requireNonNull(border, "border must not be null");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }
}
