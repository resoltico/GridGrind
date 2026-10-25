package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillInput;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.ColorInput;
import java.util.List;

/** Owns style-focused nested type groups so the main registry stays seam-sized. */
final class GridGrindProtocolCatalogStyleTypeGroups {
  static final CatalogNestedTypeDescriptor COLOR_INPUT_TYPES =
      CatalogTypeEntryFactory.nestedTypeGroup(
          "colorInputTypes",
          ColorInput.class,
          List.of(
              CatalogTypeEntryFactory.descriptor(
                  ColorInput.Rgb.class,
                  "RGB",
                  "Write one explicit RGB color reference plus optional tint metadata."),
              CatalogTypeEntryFactory.descriptor(
                  ColorInput.Theme.class,
                  "THEME",
                  "Write one workbook theme-slot color reference plus optional tint metadata."),
              CatalogTypeEntryFactory.descriptor(
                  ColorInput.Indexed.class,
                  "INDEXED",
                  "Write one indexed-palette color reference plus optional tint metadata.")));

  static final CatalogNestedTypeDescriptor CELL_GRADIENT_FILL_INPUT_TYPES =
      CatalogTypeEntryFactory.nestedTypeGroup(
          "cellGradientFillInputTypes",
          CellGradientFillInput.class,
          List.of(
              CatalogTypeEntryFactory.descriptor(
                  CellGradientFillInput.Linear.class,
                  "LINEAR",
                  "Write one linear gradient fill with ordered stops."
                      + " degree is optional when Excel's default angle is acceptable."),
              CatalogTypeEntryFactory.descriptor(
                  CellGradientFillInput.Path.class,
                  "PATH",
                  "Write one path gradient fill with ordered stops."
                      + " Each edge offset is optional and omitted offsets preserve Excel's"
                      + " default path geometry.")));

  static final CatalogNestedTypeDescriptor CELL_FILL_INPUT_TYPES =
      CatalogTypeEntryFactory.nestedTypeGroup(
          "cellFillInputTypes",
          CellFillInput.class,
          List.of(
              CatalogTypeEntryFactory.descriptor(
                  CellFillInput.PatternOnly.class,
                  "PATTERN_ONLY",
                  "Write one patterned fill with no explicit foreground or background colors."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillInput.PatternForeground.class,
                  "PATTERN_FOREGROUND",
                  "Write one patterned fill with one explicit foreground color."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillInput.PatternBackground.class,
                  "PATTERN_BACKGROUND",
                  "Write one patterned fill with one explicit background color."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillInput.PatternForegroundBackground.class,
                  "PATTERN_FOREGROUND_BACKGROUND",
                  "Write one patterned fill with explicit foreground and background colors."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillInput.Gradient.class, "GRADIENT", "Write one gradient fill patch.")));

  static final CatalogNestedTypeDescriptor CELL_COLOR_REPORT_TYPES =
      CatalogTypeEntryFactory.nestedTypeGroup(
          "cellColorReportTypes",
          CellColorReport.class,
          List.of(
              CatalogTypeEntryFactory.descriptor(
                  CellColorReport.Rgb.class,
                  "RGB",
                  "Read one explicit RGB workbook color plus optional tint metadata."),
              CatalogTypeEntryFactory.descriptor(
                  CellColorReport.Theme.class,
                  "THEME",
                  "Read one workbook theme-slot color plus optional tint metadata."),
              CatalogTypeEntryFactory.descriptor(
                  CellColorReport.Indexed.class,
                  "INDEXED",
                  "Read one indexed-palette workbook color plus optional tint metadata.")));

  static final CatalogNestedTypeDescriptor CELL_BORDER_SIDE_REPORT_TYPES =
      CatalogTypeEntryFactory.nestedTypeGroup(
          "cellBorderSideReportTypes",
          CellBorderSideReport.class,
          List.of(
              CatalogTypeEntryFactory.descriptor(
                  CellBorderSideReport.None.class,
                  "NONE",
                  "Read one border side with no visible border."),
              CatalogTypeEntryFactory.descriptor(
                  CellBorderSideReport.DefaultColor.class,
                  "DEFAULT_COLOR",
                  "Read one visible border side using Excel's implicit default color."),
              CatalogTypeEntryFactory.descriptor(
                  CellBorderSideReport.Colored.class,
                  "COLORED",
                  "Read one visible border side with an explicit color reference.")));

  static final CatalogNestedTypeDescriptor CELL_GRADIENT_FILL_REPORT_TYPES =
      CatalogTypeEntryFactory.nestedTypeGroup(
          "cellGradientFillReportTypes",
          CellGradientFillReport.class,
          List.of(
              CatalogTypeEntryFactory.descriptor(
                  CellGradientFillReport.Linear.class,
                  "LINEAR",
                  "Read one factual linear gradient fill with ordered stops."),
              CatalogTypeEntryFactory.descriptor(
                  CellGradientFillReport.Path.class,
                  "PATH",
                  "Read one factual path gradient fill with ordered stops."
                      + " Edge offsets are omitted when Excel does not persist them.")));

  static final CatalogNestedTypeDescriptor CELL_FILL_REPORT_TYPES =
      CatalogTypeEntryFactory.nestedTypeGroup(
          "cellFillReportTypes",
          CellFillReport.class,
          List.of(
              CatalogTypeEntryFactory.descriptor(
                  CellFillReport.PatternOnly.class,
                  "PATTERN_ONLY",
                  "Read one patterned fill with no explicit foreground or background colors."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillReport.PatternForeground.class,
                  "PATTERN_FOREGROUND",
                  "Read one patterned fill with one factual foreground color."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillReport.PatternBackground.class,
                  "PATTERN_BACKGROUND",
                  "Read one patterned fill with one factual background color."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillReport.PatternForegroundBackground.class,
                  "PATTERN_FOREGROUND_BACKGROUND",
                  "Read one patterned fill with factual foreground and background colors."),
              CatalogTypeEntryFactory.descriptor(
                  CellFillReport.Gradient.class, "GRADIENT", "Read one factual gradient fill.")));

  private GridGrindProtocolCatalogStyleTypeGroups() {}
}
