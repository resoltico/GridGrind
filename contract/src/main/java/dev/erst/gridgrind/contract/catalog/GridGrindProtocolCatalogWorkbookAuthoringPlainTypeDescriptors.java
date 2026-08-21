package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.BorderSideInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellGradientStopInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStylePatchInput;
import dev.erst.gridgrind.contract.dto.CommentAnchorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingDefinitionInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.IgnoredErrorInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintMarginsInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.SheetDefaultsInput;
import dev.erst.gridgrind.contract.dto.SheetDisplayInput;
import dev.erst.gridgrind.contract.dto.SheetOutlineSummaryInput;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableColumnInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.query.CellReadProjection;
import java.util.List;

/**
 * Workbook authoring plain type descriptors for comments, styles, validation, layout, and tables.
 */
@SuppressWarnings("PMD.ExcessiveImports")
final class GridGrindProtocolCatalogWorkbookAuthoringPlainTypeDescriptors {
  private GridGrindProtocolCatalogWorkbookAuthoringPlainTypeDescriptors() {}

  static final List<CatalogPlainTypeDescriptor> DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "commentInputType",
              CommentInput.class,
              "CommentInput",
              "Comment payload attached to one cell."
                  + " Comments can carry ordered rich-text runs and an explicit anchor box."),
          plainTypeDescriptor(
              "commentAnchorInputType",
              CommentAnchorInput.class,
              "CommentAnchorInput",
              "Explicit comment-anchor bounds measured in zero-based column and row indexes."),
          plainTypeDescriptor(
              "sheetProtectionSettingsType",
              SheetProtectionSettings.class,
              "SheetProtectionSettings",
              "Supported sheet-protection lock flags authored and reported by GridGrind."),
          plainTypeDescriptor(
              "cellStylePatchInputType",
              CellStylePatchInput.class,
              "CellStylePatchInput",
              "Style patch applied to a cell or range; at least one field must be set."
                  + " Colors preserve RGB, theme, indexed, and tint semantics; style subgroups"
                  + " are nested explicitly."),
          plainTypeDescriptor(
              "cellAlignmentInputType",
              CellAlignmentInput.class,
              "CellAlignmentInput",
              "Alignment patch for cell styling; at least one field must be set."
                  + " textRotation uses XSSF's explicit 0-180 degree scale and indentation uses"
                  + " Excel's 0-250 cell-indent range."),
          plainTypeDescriptor(
              "cellFontInputType",
              CellFontInput.class,
              "CellFontInput",
              "Font patch for cell styling; at least one field must be set."
                  + " Colors can use RGB, theme, indexed, and tint semantics."),
          plainTypeDescriptor(
              "richTextRunInputType",
              RichTextRunInput.class,
              "RichTextRunInput",
              "One ordered rich-text run for a string cell."
                  + " text must be non-empty; font is an optional override patch."
                  + " The ordered run texts concatenate to the stored plain string value."),
          plainTypeDescriptor(
              "cellReadProjectionType",
              CellReadProjection.class,
              "CellReadProjection",
              "Shared readback projection that selects which factual cell facets are included on"
                  + " cell-returning inspection surfaces."
                  + " Facets default to VALUE when omitted by the enclosing query."
                  + " The facets field publishes enumValueDocs so each facet token explains the"
                  + " response fields it unlocks."
                  + " Catalog field descriptors publish facet-gated response fields through"
                  + " projectedByFacets so agents can derive the required facet directly from the"
                  + " machine contract."
                  + " Date-like numeric cells still read back as type=NUMBER until TEMPORAL is"
                  + " requested."),
          plainTypeDescriptor(
              "cellGradientStopInputType",
              CellGradientStopInput.class,
              "CellGradientStopInput",
              "One gradient stop with a normalized position between 0.0 and 1.0."),
          plainTypeDescriptor(
              "cellBorderInputType",
              CellBorderInput.class,
              "CellBorderInput",
              "Border patch for cell styling; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides."),
          plainTypeDescriptor(
              "borderSideInputType",
              BorderSideInput.class,
              "BorderSideInput",
              "One border side shared by cell and differential style patches, defined by its"
                  + " border style and optional color semantics."),
          plainTypeDescriptor(
              "cellProtectionInputType",
              CellProtectionInput.class,
              "CellProtectionInput",
              "Cell protection patch; at least one field must be set."
                  + " These flags matter when sheet protection is enabled."),
          plainTypeDescriptor(
              "dataValidationInputType",
              DataValidationInput.class,
              "DataValidationInput",
              "Supported data-validation definition attached to one sheet range."),
          plainTypeDescriptor(
              "dataValidationPromptInputType",
              DataValidationPromptInput.class,
              "DataValidationPromptInput",
              "Optional prompt-box configuration shown when a validated cell is selected."),
          plainTypeDescriptor(
              "dataValidationErrorAlertInputType",
              DataValidationErrorAlertInput.class,
              "DataValidationErrorAlertInput",
              "Optional error-box configuration shown when invalid data is entered."),
          plainTypeDescriptor(
              "autofilterCustomConditionInputType",
              AutofilterFilterCriterionInput.CustomConditionInput.class,
              "AutofilterCustomConditionInput",
              "One comparator-value pair nested inside a custom autofilter criterion."),
          plainTypeDescriptor(
              "autofilterFilterColumnInputType",
              AutofilterFilterColumnInput.class,
              "AutofilterFilterColumnInput",
              "One authored autofilter filter-column payload with an explicit column criterion."),
          plainTypeDescriptor(
              "autofilterSortStateInputType",
              AutofilterSortStateInput.class,
              "AutofilterSortStateInput",
              "Authored autofilter sort-state payload with one or more ordered sort conditions."),
          plainTypeDescriptor(
              "conditionalFormattingDefinitionInputType",
              ConditionalFormattingDefinitionInput.class,
              "ConditionalFormattingDefinitionInput",
              "One authored conditional-formatting definition with an ordered rule list."
                  + " Target ranges are owned by the selector; rules must not be empty."),
          plainTypeDescriptor(
              "conditionalFormattingThresholdInputType",
              ConditionalFormattingThresholdInput.class,
              "ConditionalFormattingThresholdInput",
              "Threshold payload shared by authored advanced conditional-formatting rules."),
          plainTypeDescriptor(
              "headerFooterTextInputType",
              HeaderFooterTextInput.class,
              "HeaderFooterTextInput",
              "Plain left, center, and right header or footer text segments."
                  + " Supply all three fields explicitly."),
          plainTypeDescriptor(
              "differentialStyleInputType",
              DifferentialStyleInput.class,
              "DifferentialStyleInput",
              "Differential style payload used by authored conditional-formatting rules."
                  + " At least one field must be set. Colors preserve RGB, theme, indexed,"
                  + " and tint semantics."),
          plainTypeDescriptor(
              "differentialBorderInputType",
              DifferentialBorderInput.class,
              "DifferentialBorderInput",
              "Conditional-formatting differential border patch; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides."),
          plainTypeDescriptor(
              "ignoredErrorInputType",
              IgnoredErrorInput.class,
              "IgnoredErrorInput",
              "One ignored-error block anchored to one A1-style range plus one or more"
                  + " ignored-error families."),
          plainTypeDescriptor(
              "printLayoutInputType",
              PrintLayoutInput.class,
              "PrintLayoutInput",
              "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request."
                  + " Supply explicit printArea, orientation, scaling, repeatingRows,"
                  + " repeatingColumns, header, footer, and setup fields."),
          plainTypeDescriptor(
              "printMarginsInputType",
              PrintMarginsInput.class,
              "PrintMarginsInput",
              "Explicit print margins measured in the workbook's stored inch-based values."),
          plainTypeDescriptor(
              "printSetupInputType",
              PrintSetupInput.class,
              "PrintSetupInput",
              "Advanced page-setup payload nested under print-layout authoring."
                  + " Supply the full authored setup block explicitly."),
          plainTypeDescriptor(
              "sheetDefaultsInputType",
              SheetDefaultsInput.class,
              "SheetDefaultsInput",
              "Default row and column sizing authored as part of sheet-presentation state."
                  + " Supply explicit defaultColumnWidth and defaultRowHeightPoints values."
                  + " defaultColumnWidth must be > 0 and <= 255;"
                  + " defaultRowHeightPoints must be > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + "."),
          plainTypeDescriptor(
              "sheetDisplayInputType",
              SheetDisplayInput.class,
              "SheetDisplayInput",
              "Screen-facing sheet display flags authored as part of sheet-presentation state."
                  + " Supply all display flags explicitly."),
          plainTypeDescriptor(
              "sheetOutlineSummaryInputType",
              SheetOutlineSummaryInput.class,
              "SheetOutlineSummaryInput",
              "Outline-summary placement authored as part of sheet-presentation state."
                  + " Supply both placement flags explicitly."),
          plainTypeDescriptor(
              "sheetPresentationInputType",
              SheetPresentationInput.class,
              "SheetPresentationInput",
              "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request."
                  + " Supply explicit display, tabColor, outlineSummary, sheetDefaults, and"
                  + " ignoredErrors values."),
          plainTypeDescriptor(
              "pivotTableInputType",
              PivotTableInput.class,
              "PivotTableInput",
              "Workbook-global pivot-table definition for one SET_PIVOT_TABLE request."
                  + " Source-column assignments across rowLabels, columnLabels, reportFilters,"
                  + " and dataFields must be disjoint."
                  + " reportFilters require anchor.topLeftAddress on row 3 or lower."),
          plainTypeDescriptor(
              "pivotTableAnchorInputType",
              PivotTableInput.Anchor.class,
              "PivotTableAnchorInput",
              "Top-left anchor for a pivot table rendered on its destination sheet."
                  + " The address must be a single-cell A1 reference."),
          plainTypeDescriptor(
              "pivotTableDataFieldInputType",
              PivotTableInput.DataField.class,
              "PivotTableDataFieldInput",
              "One authored pivot data field bound to a source column and aggregation function."
                  + " displayName defaults to sourceColumnName when omitted."),
          plainTypeDescriptor(
              "tableColumnInputType",
              TableColumnInput.class,
              "TableColumnInput",
              "Advanced table-column metadata applied by zero-based ordinal column index."),
          plainTypeDescriptor(
              "tableInputType",
              TableInput.class,
              "TableInput",
              "Workbook-global table definition for one SET_TABLE request."),
          plainTypeDescriptor(
              "workbookProtectionInputType",
              WorkbookProtectionInput.class,
              "WorkbookProtectionInput",
              "Workbook-protection payload covering workbook and revisions lock state plus"
                  + " optional passwords."));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group, Class<? extends Record> recordType, String id, String summary) {
    return CatalogTypeEntryFactory.plainTypeDescriptor(group, recordType, id, summary);
  }
}
