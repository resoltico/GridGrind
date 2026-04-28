package dev.erst.gridgrind.contract.catalog;

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
      GridGrindProtocolCatalog.nestedTypeGroup(
          "colorInputTypes",
          ColorInput.class,
          List.of(
              GridGrindProtocolCatalog.descriptor(
                  ColorInput.Rgb.class,
                  "RGB",
                  "Write one explicit RGB color reference plus optional tint metadata.",
                  "tint"),
              GridGrindProtocolCatalog.descriptor(
                  ColorInput.Theme.class,
                  "THEME",
                  "Write one workbook theme-slot color reference plus optional tint metadata.",
                  "tint"),
              GridGrindProtocolCatalog.descriptor(
                  ColorInput.Indexed.class,
                  "INDEXED",
                  "Write one indexed-palette color reference plus optional tint metadata.",
                  "tint")));

  static final CatalogNestedTypeDescriptor CELL_GRADIENT_FILL_INPUT_TYPES =
      GridGrindProtocolCatalog.nestedTypeGroup(
          "cellGradientFillInputTypes",
          CellGradientFillInput.class,
          List.of(
              GridGrindProtocolCatalog.descriptor(
                  CellGradientFillInput.Linear.class,
                  "LINEAR",
                  "Write one linear gradient fill with ordered stops."
                      + " degree is optional when Excel's default angle is acceptable.",
                  "degree"),
              GridGrindProtocolCatalog.descriptor(
                  CellGradientFillInput.Path.class,
                  "PATH",
                  "Write one path gradient fill with ordered stops."
                      + " Each edge offset is optional and omitted offsets preserve Excel's"
                      + " default path geometry.",
                  "left",
                  "right",
                  "top",
                  "bottom")));

  static final CatalogNestedTypeDescriptor CELL_FILL_INPUT_TYPES =
      GridGrindProtocolCatalog.nestedTypeGroup(
          "cellFillInputTypes",
          CellFillInput.class,
          List.of(
              GridGrindProtocolCatalog.descriptor(
                  CellFillInput.PatternOnly.class,
                  "PATTERN_ONLY",
                  "Write one patterned fill with no explicit foreground or background colors."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillInput.PatternForeground.class,
                  "PATTERN_FOREGROUND",
                  "Write one patterned fill with one explicit foreground color."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillInput.PatternBackground.class,
                  "PATTERN_BACKGROUND",
                  "Write one patterned fill with one explicit background color."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillInput.PatternForegroundBackground.class,
                  "PATTERN_FOREGROUND_BACKGROUND",
                  "Write one patterned fill with explicit foreground and background colors."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillInput.Gradient.class, "GRADIENT", "Write one gradient fill patch.")));

  static final CatalogNestedTypeDescriptor CELL_COLOR_REPORT_TYPES =
      GridGrindProtocolCatalog.nestedTypeGroup(
          "cellColorReportTypes",
          CellColorReport.class,
          List.of(
              GridGrindProtocolCatalog.descriptor(
                  CellColorReport.Rgb.class,
                  "RGB",
                  "Read one explicit RGB workbook color plus optional tint metadata.",
                  "tint"),
              GridGrindProtocolCatalog.descriptor(
                  CellColorReport.Theme.class,
                  "THEME",
                  "Read one workbook theme-slot color plus optional tint metadata.",
                  "tint"),
              GridGrindProtocolCatalog.descriptor(
                  CellColorReport.Indexed.class,
                  "INDEXED",
                  "Read one indexed-palette workbook color plus optional tint metadata.",
                  "tint")));

  static final CatalogNestedTypeDescriptor CELL_GRADIENT_FILL_REPORT_TYPES =
      GridGrindProtocolCatalog.nestedTypeGroup(
          "cellGradientFillReportTypes",
          CellGradientFillReport.class,
          List.of(
              GridGrindProtocolCatalog.descriptor(
                  CellGradientFillReport.Linear.class,
                  "LINEAR",
                  "Read one factual linear gradient fill with ordered stops.",
                  "degree"),
              GridGrindProtocolCatalog.descriptor(
                  CellGradientFillReport.Path.class,
                  "PATH",
                  "Read one factual path gradient fill with ordered stops."
                      + " Edge offsets are omitted when Excel does not persist them.",
                  "left",
                  "right",
                  "top",
                  "bottom")));

  static final CatalogNestedTypeDescriptor CELL_FILL_REPORT_TYPES =
      GridGrindProtocolCatalog.nestedTypeGroup(
          "cellFillReportTypes",
          CellFillReport.class,
          List.of(
              GridGrindProtocolCatalog.descriptor(
                  CellFillReport.PatternOnly.class,
                  "PATTERN_ONLY",
                  "Read one patterned fill with no explicit foreground or background colors."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillReport.PatternForeground.class,
                  "PATTERN_FOREGROUND",
                  "Read one patterned fill with one factual foreground color."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillReport.PatternBackground.class,
                  "PATTERN_BACKGROUND",
                  "Read one patterned fill with one factual background color."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillReport.PatternForegroundBackground.class,
                  "PATTERN_FOREGROUND_BACKGROUND",
                  "Read one patterned fill with factual foreground and background colors."),
              GridGrindProtocolCatalog.descriptor(
                  CellFillReport.Gradient.class, "GRADIENT", "Read one factual gradient fill.")));

  private GridGrindProtocolCatalogStyleTypeGroups() {}
}
