package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelColor;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Converts workbook style snapshots into protocol style report records. */
final class InspectionResultCellStyleReportSupport {
  private InspectionResultCellStyleReportSupport() {}

  static Optional<CellColorReport> toCellColorReport(@Nullable ExcelColor color) {
    return color == null
        ? Optional.empty()
        : Optional.of(
            switch (color) {
              case ExcelColor.Rgb rgb ->
                  rgb.tint().isPresent()
                      ? CellColorReport.rgb(new RgbHex(rgb.rgb()), rgb.tint().orElseThrow())
                      : CellColorReport.rgb(new RgbHex(rgb.rgb()));
              case ExcelColor.Theme theme ->
                  theme.tint().isPresent()
                      ? CellColorReport.theme(theme.theme(), theme.tint().orElseThrow())
                      : CellColorReport.theme(theme.theme());
              case ExcelColor.Indexed indexed ->
                  indexed.tint().isPresent()
                      ? CellColorReport.indexed(indexed.indexed(), indexed.tint().orElseThrow())
                      : CellColorReport.indexed(indexed.indexed());
            });
  }
}
