package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellGradientStopInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.CommentAnchorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.IgnoredErrorInput;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
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
                  + " Comments can carry ordered rich-text runs and an explicit anchor box.",
              List.of("visible", "runs", "anchor")),
          plainTypeDescriptor(
              "commentAnchorInputType",
              CommentAnchorInput.class,
              "CommentAnchorInput",
              "Explicit comment-anchor bounds measured in zero-based column and row indexes.",
              List.of()),
          plainTypeDescriptor(
              "namedRangeTargetType",
              NamedRangeTarget.class,
              "NamedRangeTarget",
              "Named-range target payload."
                  + " Supply either sheetName plus range, or formula by itself.",
              List.of("sheetName", "range", "formula")),
          plainTypeDescriptor(
              "sheetProtectionSettingsType",
              SheetProtectionSettings.class,
              "SheetProtectionSettings",
              "Supported sheet-protection lock flags authored and reported by GridGrind.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleInputType",
              CellStyleInput.class,
              "CellStyleInput",
              "Style patch applied to a cell or range; at least one field must be set."
                  + " Colors use #RRGGBB hex and style subgroups are nested explicitly.",
              List.of("numberFormat", "alignment", "font", "fill", "border", "protection")),
          plainTypeDescriptor(
              "cellAlignmentInputType",
              CellAlignmentInput.class,
              "CellAlignmentInput",
              "Alignment patch for cell styling; at least one field must be set."
                  + " textRotation uses XSSF's explicit 0-180 degree scale and indentation uses"
                  + " Excel's 0-250 cell-indent range.",
              List.of(
                  "wrapText",
                  "horizontalAlignment",
                  "verticalAlignment",
                  "textRotation",
                  "indentation")),
          plainTypeDescriptor(
              "cellFontInputType",
              CellFontInput.class,
              "CellFontInput",
              "Font patch for cell styling; at least one field must be set."
                  + " Colors can use RGB, theme, indexed, and tint semantics.",
              List.of(
                  "bold",
                  "italic",
                  "fontName",
                  "fontHeight",
                  "fontColor",
                  "underline",
                  "strikeout")),
          plainTypeDescriptor(
              "richTextRunInputType",
              RichTextRunInput.class,
              "RichTextRunInput",
              "One ordered rich-text run for a string cell."
                  + " text must be non-empty; font is an optional override patch."
                  + " The ordered run texts concatenate to the stored plain string value.",
              List.of("font")),
          plainTypeDescriptor(
              "cellGradientStopInputType",
              CellGradientStopInput.class,
              "CellGradientStopInput",
              "One gradient stop with a normalized position between 0.0 and 1.0.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderInputType",
              CellBorderInput.class,
              "CellBorderInput",
              "Border patch for cell styling; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "cellBorderSideInputType",
              CellBorderSideInput.class,
              "CellBorderSideInput",
              "One border side defined by its border style and optional color semantics.",
              List.of("style", "color")),
          plainTypeDescriptor(
              "cellProtectionInputType",
              CellProtectionInput.class,
              "CellProtectionInput",
              "Cell protection patch; at least one field must be set."
                  + " These flags matter when sheet protection is enabled.",
              List.of("locked", "hiddenFormula")),
          plainTypeDescriptor(
              "dataValidationInputType",
              DataValidationInput.class,
              "DataValidationInput",
              "Supported data-validation definition attached to one sheet range.",
              List.of("allowBlank", "suppressDropDownArrow", "prompt", "errorAlert")),
          plainTypeDescriptor(
              "dataValidationPromptInputType",
              DataValidationPromptInput.class,
              "DataValidationPromptInput",
              "Optional prompt-box configuration shown when a validated cell is selected.",
              List.of("showPromptBox")),
          plainTypeDescriptor(
              "dataValidationErrorAlertInputType",
              DataValidationErrorAlertInput.class,
              "DataValidationErrorAlertInput",
              "Optional error-box configuration shown when invalid data is entered.",
              List.of("showErrorBox")),
          plainTypeDescriptor(
              "autofilterCustomConditionInputType",
              AutofilterFilterCriterionInput.CustomConditionInput.class,
              "AutofilterCustomConditionInput",
              "One comparator-value pair nested inside a custom autofilter criterion.",
              List.of()),
          plainTypeDescriptor(
              "autofilterFilterColumnInputType",
              AutofilterFilterColumnInput.class,
              "AutofilterFilterColumnInput",
              "One authored autofilter filter-column payload with an explicit column criterion.",
              List.of("showButton")),
          plainTypeDescriptor(
              "autofilterSortStateInputType",
              AutofilterSortStateInput.class,
              "AutofilterSortStateInput",
              "Authored autofilter sort-state payload with one or more ordered sort conditions.",
              List.of("caseSensitive", "columnSort", "sortMethod")),
          plainTypeDescriptor(
              "conditionalFormattingBlockInputType",
              ConditionalFormattingBlockInput.class,
              "ConditionalFormattingBlockInput",
              "One authored conditional-formatting block with ordered target ranges and rules."
                  + " rules must not be empty; ranges must be unique.",
              List.of()),
          plainTypeDescriptor(
              "conditionalFormattingThresholdInputType",
              ConditionalFormattingThresholdInput.class,
              "ConditionalFormattingThresholdInput",
              "Threshold payload shared by authored advanced conditional-formatting rules.",
              List.of("formula", "value")),
          plainTypeDescriptor(
              "headerFooterTextInputType",
              HeaderFooterTextInput.class,
              "HeaderFooterTextInput",
              "Plain left, center, and right header or footer text segments."
                  + " Supply all three fields explicitly.",
              List.of()),
          plainTypeDescriptor(
              "differentialStyleInputType",
              DifferentialStyleInput.class,
              "DifferentialStyleInput",
              "Differential style payload used by authored conditional-formatting rules."
                  + " At least one field must be set. Colors use #RRGGBB hex.",
              List.of(
                  "numberFormat",
                  "bold",
                  "italic",
                  "fontHeight",
                  "fontColor",
                  "underline",
                  "strikeout",
                  "fillColor",
                  "border")),
          plainTypeDescriptor(
              "differentialBorderInputType",
              DifferentialBorderInput.class,
              "DifferentialBorderInput",
              "Conditional-formatting differential border patch; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "differentialBorderSideInputType",
              DifferentialBorderSideInput.class,
              "DifferentialBorderSideInput",
              "One conditional-formatting differential border side defined by style and optional"
                  + " color.",
              List.of()),
          plainTypeDescriptor(
              "ignoredErrorInputType",
              IgnoredErrorInput.class,
              "IgnoredErrorInput",
              "One ignored-error block anchored to one A1-style range plus one or more"
                  + " ignored-error families.",
              List.of()),
          plainTypeDescriptor(
              "printLayoutInputType",
              PrintLayoutInput.class,
              "PrintLayoutInput",
              "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request."
                  + " Supply explicit printArea, orientation, scaling, repeatingRows,"
                  + " repeatingColumns, header, footer, and setup fields.",
              List.of()),
          plainTypeDescriptor(
              "printMarginsInputType",
              PrintMarginsInput.class,
              "PrintMarginsInput",
              "Explicit print margins measured in the workbook's stored inch-based values.",
              List.of()),
          plainTypeDescriptor(
              "printSetupInputType",
              PrintSetupInput.class,
              "PrintSetupInput",
              "Advanced page-setup payload nested under print-layout authoring."
                  + " Supply the full authored setup block explicitly.",
              List.of()),
          plainTypeDescriptor(
              "sheetDefaultsInputType",
              SheetDefaultsInput.class,
              "SheetDefaultsInput",
              "Default row and column sizing authored as part of sheet-presentation state."
                  + " Supply explicit defaultColumnWidth and defaultRowHeightPoints values."
                  + " defaultColumnWidth must be > 0 and <= 255;"
                  + " defaultRowHeightPoints must be > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + ".",
              List.of()),
          plainTypeDescriptor(
              "sheetDisplayInputType",
              SheetDisplayInput.class,
              "SheetDisplayInput",
              "Screen-facing sheet display flags authored as part of sheet-presentation state."
                  + " Supply all display flags explicitly.",
              List.of()),
          plainTypeDescriptor(
              "sheetOutlineSummaryInputType",
              SheetOutlineSummaryInput.class,
              "SheetOutlineSummaryInput",
              "Outline-summary placement authored as part of sheet-presentation state."
                  + " Supply both placement flags explicitly.",
              List.of()),
          plainTypeDescriptor(
              "sheetPresentationInputType",
              SheetPresentationInput.class,
              "SheetPresentationInput",
              "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request."
                  + " Supply explicit display, tabColor, outlineSummary, sheetDefaults, and"
                  + " ignoredErrors values.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableInputType",
              PivotTableInput.class,
              "PivotTableInput",
              "Workbook-global pivot-table definition for one SET_PIVOT_TABLE request."
                  + " Source-column assignments across rowLabels, columnLabels, reportFilters,"
                  + " and dataFields must be disjoint."
                  + " reportFilters require anchor.topLeftAddress on row 3 or lower.",
              List.of("rowLabels", "columnLabels", "reportFilters")),
          plainTypeDescriptor(
              "pivotTableAnchorInputType",
              PivotTableInput.Anchor.class,
              "PivotTableAnchorInput",
              "Top-left anchor for a pivot table rendered on its destination sheet."
                  + " The address must be a single-cell A1 reference.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldInputType",
              PivotTableInput.DataField.class,
              "PivotTableDataFieldInput",
              "One authored pivot data field bound to a source column and aggregation function."
                  + " displayName defaults to sourceColumnName when omitted.",
              List.of("displayName", "valueFormat")),
          plainTypeDescriptor(
              "tableColumnInputType",
              TableColumnInput.class,
              "TableColumnInput",
              "Advanced table-column metadata applied by zero-based ordinal column index.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "tableInputType",
              TableInput.class,
              "TableInput",
              "Workbook-global table definition for one SET_TABLE request.",
              List.of(
                  "showTotalsRow",
                  "hasAutofilter",
                  "comment",
                  "published",
                  "insertRow",
                  "insertRowShift",
                  "headerRowCellStyle",
                  "dataCellStyle",
                  "totalsRowCellStyle",
                  "columns")),
          plainTypeDescriptor(
              "workbookProtectionInputType",
              WorkbookProtectionInput.class,
              "WorkbookProtectionInput",
              "Workbook-protection payload covering workbook and revisions lock state plus"
                  + " optional passwords.",
              List.of(
                  "structureLocked",
                  "windowsLocked",
                  "revisionsLocked",
                  "workbookPassword",
                  "revisionsPassword")));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return GridGrindProtocolCatalog.plainTypeDescriptor(
        group, recordType, id, summary, optionalFields);
  }
}
