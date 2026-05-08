package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.CellCommentReport;
import dev.erst.gridgrind.contract.dto.CellHyperlinkReport;
import dev.erst.gridgrind.contract.dto.CellReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.contract.dto.DataValidationEntryReport;
import dev.erst.gridgrind.contract.dto.MergedRegionReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.SheetLayoutReport;
import dev.erst.gridgrind.contract.dto.SheetSummaryReport;
import dev.erst.gridgrind.contract.dto.WindowReport;
import java.util.List;
import java.util.Objects;

/** Sheet-scoped factual inspection results. */
public sealed interface SheetInspectionResult extends InspectionIntrospectionResult
    permits SheetInspectionResult.SheetSummaryResult,
        SheetInspectionResult.ArrayFormulasResult,
        SheetInspectionResult.CellsResult,
        SheetInspectionResult.WindowResult,
        SheetInspectionResult.MergedRegionsResult,
        SheetInspectionResult.HyperlinksResult,
        SheetInspectionResult.CommentsResult,
        SheetInspectionResult.SheetLayoutResult,
        SheetInspectionResult.PrintLayoutResult,
        SheetInspectionResult.DataValidationsResult,
        SheetInspectionResult.ConditionalFormattingResult,
        SheetInspectionResult.AutofiltersResult {

  /** Returns summary facts for one sheet. */
  @JsonTypeName("GET_SHEET_SUMMARY")
  record SheetSummaryResult(String stepId, SheetSummaryReport sheet)
      implements SheetInspectionResult {
    public SheetSummaryResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Returns factual array-formula groups across the selected sheets. */
  @JsonTypeName("GET_ARRAY_FORMULAS")
  record ArrayFormulasResult(String stepId, List<ArrayFormulaReport> arrayFormulas)
      implements SheetInspectionResult {
    public ArrayFormulasResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      arrayFormulas = InspectionResultValidationSupport.copyValues(arrayFormulas, "arrayFormulas");
    }
  }

  /** Returns exact cell snapshots for one sheet. */
  @JsonTypeName("GET_CELLS")
  record CellsResult(String stepId, String sheetName, List<CellReport> cells)
      implements SheetInspectionResult {
    public CellsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      cells = InspectionResultValidationSupport.copyValues(cells, "cells");
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  @JsonTypeName("GET_WINDOW")
  record WindowResult(String stepId, WindowReport window) implements SheetInspectionResult {
    public WindowResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(window, "window must not be null");
    }
  }

  /** Returns every merged region present on one sheet. */
  @JsonTypeName("GET_MERGED_REGIONS")
  record MergedRegionsResult(
      String stepId, String sheetName, List<MergedRegionReport> mergedRegions)
      implements SheetInspectionResult {
    public MergedRegionsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      mergedRegions = InspectionResultValidationSupport.copyValues(mergedRegions, "mergedRegions");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  @JsonTypeName("GET_HYPERLINKS")
  record HyperlinksResult(String stepId, String sheetName, List<CellHyperlinkReport> hyperlinks)
      implements SheetInspectionResult {
    public HyperlinksResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      hyperlinks = InspectionResultValidationSupport.copyValues(hyperlinks, "hyperlinks");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  @JsonTypeName("GET_COMMENTS")
  record CommentsResult(String stepId, String sheetName, List<CellCommentReport> comments)
      implements SheetInspectionResult {
    public CommentsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      comments = InspectionResultValidationSupport.copyValues(comments, "comments");
    }
  }

  /** Returns layout facts such as pane state, zoom, and explicit sizing. */
  @JsonTypeName("GET_SHEET_LAYOUT")
  record SheetLayoutResult(String stepId, SheetLayoutReport layout)
      implements SheetInspectionResult {
    public SheetLayoutResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns supported print-layout facts for one sheet. */
  @JsonTypeName("GET_PRINT_LAYOUT")
  record PrintLayoutResult(String stepId, PrintLayoutReport layout)
      implements SheetInspectionResult {
    public PrintLayoutResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  @JsonTypeName("GET_DATA_VALIDATIONS")
  record DataValidationsResult(
      String stepId, String sheetName, List<DataValidationEntryReport> validations)
      implements SheetInspectionResult {
    public DataValidationsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      validations = InspectionResultValidationSupport.copyValues(validations, "validations");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  @JsonTypeName("GET_CONDITIONAL_FORMATTING")
  record ConditionalFormattingResult(
      String stepId,
      String sheetName,
      List<ConditionalFormattingEntryReport> conditionalFormattingBlocks)
      implements SheetInspectionResult {
    public ConditionalFormattingResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      conditionalFormattingBlocks =
          InspectionResultValidationSupport.copyValues(
              conditionalFormattingBlocks, "conditionalFormattingBlocks");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  @JsonTypeName("GET_AUTOFILTERS")
  record AutofiltersResult(String stepId, String sheetName, List<AutofilterEntryReport> autofilters)
      implements SheetInspectionResult {
    public AutofiltersResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      autofilters = InspectionResultValidationSupport.copyValues(autofilters, "autofilters");
    }
  }
}
