package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
  record SheetSummaryResult(StepId stepId, SheetSummaryReport sheet)
      implements SheetInspectionResult {
    public SheetSummaryResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Returns factual array-formula groups across the selected sheets. */
  @JsonTypeName("GET_ARRAY_FORMULAS")
  record ArrayFormulasResult(StepId stepId, List<ArrayFormulaReport> arrayFormulas)
      implements SheetInspectionResult {
    public ArrayFormulasResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      arrayFormulas = InspectionResultValidationSupport.copyValues(arrayFormulas, "arrayFormulas");
    }
  }

  /** Returns exact cell snapshots for one sheet. */
  @JsonTypeName("GET_CELLS")
  record CellsResult(StepId stepId, SheetLabel sheetName, List<CellReport> cells)
      implements SheetInspectionResult {
    public CellsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      cells = InspectionResultValidationSupport.copyValues(cells, "cells");
    }
  }

  /** Returns a rectangular window of cell snapshots anchored at one top-left address. */
  @JsonTypeName("GET_WINDOW")
  record WindowResult(StepId stepId, WindowReport window) implements SheetInspectionResult {
    public WindowResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(window, "window must not be null");
    }
  }

  /** Returns every merged region present on one sheet. */
  @JsonTypeName("GET_MERGED_REGIONS")
  record MergedRegionsResult(
      StepId stepId, SheetLabel sheetName, List<MergedRegionReport> mergedRegions)
      implements SheetInspectionResult {
    public MergedRegionsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      mergedRegions = InspectionResultValidationSupport.copyValues(mergedRegions, "mergedRegions");
    }
  }

  /** Returns hyperlink metadata for selected cells on one sheet. */
  @JsonTypeName("GET_HYPERLINKS")
  record HyperlinksResult(StepId stepId, SheetLabel sheetName, List<CellHyperlinkReport> hyperlinks)
      implements SheetInspectionResult {
    public HyperlinksResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      hyperlinks = InspectionResultValidationSupport.copyValues(hyperlinks, "hyperlinks");
    }
  }

  /** Returns comment metadata for selected cells on one sheet. */
  @JsonTypeName("GET_COMMENTS")
  record CommentsResult(StepId stepId, SheetLabel sheetName, List<CellCommentReport> comments)
      implements SheetInspectionResult {
    public CommentsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      comments = InspectionResultValidationSupport.copyValues(comments, "comments");
    }
  }

  /** Returns layout facts such as pane state, zoom, and explicit sizing. */
  @JsonTypeName("GET_SHEET_LAYOUT")
  record SheetLayoutResult(StepId stepId, SheetLayoutReport layout)
      implements SheetInspectionResult {
    public SheetLayoutResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns supported print-layout facts for one sheet. */
  @JsonTypeName("GET_PRINT_LAYOUT")
  record PrintLayoutResult(StepId stepId, PrintLayoutReport layout)
      implements SheetInspectionResult {
    public PrintLayoutResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      Objects.requireNonNull(layout, "layout must not be null");
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  @JsonTypeName("GET_DATA_VALIDATIONS")
  record DataValidationsResult(
      StepId stepId, SheetLabel sheetName, List<DataValidationEntryReport> validations)
      implements SheetInspectionResult {
    public DataValidationsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      validations = InspectionResultValidationSupport.copyValues(validations, "validations");
    }
  }

  /** Returns conditional-formatting metadata for the selected ranges on one sheet. */
  @JsonTypeName("GET_CONDITIONAL_FORMATTING")
  record ConditionalFormattingResult(
      StepId stepId,
      SheetLabel sheetName,
      List<ConditionalFormattingEntryReport> conditionalFormattingBlocks)
      implements SheetInspectionResult {
    public ConditionalFormattingResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      conditionalFormattingBlocks =
          InspectionResultValidationSupport.copyValues(
              conditionalFormattingBlocks, "conditionalFormattingBlocks");
    }
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  @JsonTypeName("GET_AUTOFILTERS")
  record AutofiltersResult(
      StepId stepId, SheetLabel sheetName, List<AutofilterEntryReport> autofilters)
      implements SheetInspectionResult {
    public AutofiltersResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      autofilters = InspectionResultValidationSupport.copyValues(autofilters, "autofilters");
    }
  }
}
