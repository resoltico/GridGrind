package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;

import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import java.util.List;

/** Owns one focused subset of nested-type group descriptors for the protocol catalog. */
final class GridGrindProtocolCatalogWorkbookInputNestedTypeGroups {
  private GridGrindProtocolCatalogWorkbookInputNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> WORKBOOK_INPUT_GROUPS =
      List.of(
          nestedTypeGroup(
              "cellInputTypes",
              CellInput.class,
              List.of(
                  descriptor(CellInput.Blank.class, "BLANK", "Write an empty cell."),
                  descriptor(
                      CellInput.Text.class,
                      "TEXT",
                      "Write a string cell value. Blank text is rejected; use BLANK for empty"
                          + " cells."),
                  descriptor(
                      CellInput.RichText.class,
                      "RICH_TEXT",
                      "Write a structured string cell value with an ordered rich-text run list."
                          + " Run text concatenates to the stored plain string value and each run"
                          + " may override font attributes independently."),
                  descriptor(CellInput.Numeric.class, "NUMBER", "Write a numeric cell value."),
                  descriptor(
                      CellInput.BooleanValue.class, "BOOLEAN", "Write a boolean cell value."),
                  descriptor(
                      CellInput.Date.class,
                      "DATE",
                      "Write an ISO-8601 date value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.DateTime.class,
                      "DATE_TIME",
                      "Write an ISO-8601 date-time value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.Formula.class,
                      "FORMULA",
                      "Write an Excel formula. A leading = sign is accepted and stripped automatically; the engine stores the formula without it."))),
          nestedTypeGroup(
              "hyperlinkTargetTypes",
              HyperlinkTarget.class,
              List.of(
                  descriptor(HyperlinkTarget.Url.class, "URL", "Attach an absolute URL target."),
                  descriptor(
                      HyperlinkTarget.Email.class,
                      "EMAIL",
                      "Attach an email target without the mailto: prefix."),
                  descriptor(
                      HyperlinkTarget.File.class,
                      "FILE",
                      "Attach a local or shared file path."
                          + " Accepts plain paths or file: URIs and normalizes to a plain path"
                          + " string."),
                  descriptor(
                      HyperlinkTarget.Document.class,
                      "DOCUMENT",
                      "Attach an internal workbook target."))),
          nestedTypeGroup(
              "paneTypes",
              PaneInput.class,
              List.of(
                  descriptor(
                      PaneInput.None.class, "NONE", "Sheet has no active pane split or freeze."),
                  descriptor(
                      PaneInput.Frozen.class,
                      "FROZEN",
                      "Freeze the sheet at the provided split and visible-origin coordinates."),
                  descriptor(
                      PaneInput.Split.class,
                      "SPLIT",
                      "Apply split panes with explicit split offsets, visible origin, and active pane."))),
          nestedTypeGroup(
              "sheetCopyPositionTypes",
              SheetCopyPosition.class,
              List.of(
                  descriptor(
                      SheetCopyPosition.AppendAtEnd.class,
                      "APPEND_AT_END",
                      "Place the copied sheet after every existing sheet."),
                  descriptor(
                      SheetCopyPosition.AtIndex.class,
                      "AT_INDEX",
                      "Place the copied sheet at the requested zero-based workbook position."))),
          nestedTypeGroup(
              "autofilterSortConditionInputTypes",
              AutofilterSortConditionInput.class,
              List.of(
                  descriptor(
                      AutofilterSortConditionInput.Value.class,
                      "VALUE",
                      "Sort by the ordinary cell value with no auxiliary discriminator payload."),
                  descriptor(
                      AutofilterSortConditionInput.CellColor.class,
                      "CELL_COLOR",
                      "Sort by the cell fill color referenced by one explicit color."),
                  descriptor(
                      AutofilterSortConditionInput.FontColor.class,
                      "FONT_COLOR",
                      "Sort by the rendered font color referenced by one explicit color."),
                  descriptor(
                      AutofilterSortConditionInput.Icon.class,
                      "ICON",
                      "Sort by the icon id inside the active icon-set definition."))));
}
