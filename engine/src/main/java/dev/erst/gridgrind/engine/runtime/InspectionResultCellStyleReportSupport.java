package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Converts workbook style snapshots into protocol style report records. */
final class InspectionResultCellStyleReportSupport {
  private InspectionResultCellStyleReportSupport() {}

  static Optional<CellColorReport> toCellColorReport(@Nullable ColorSpec color) {
    return color == null
        ? Optional.empty()
        : Optional.of(
            switch (color) {
              case ColorSpec.Rgb rgb ->
                  rgb.tint().isPresent()
                      ? CellColorReport.rgb(rgb.rgb(), rgb.tint().orElseThrow())
                      : CellColorReport.rgb(rgb.rgb());
              case ColorSpec.Theme theme ->
                  theme.tint().isPresent()
                      ? CellColorReport.theme(theme.theme(), theme.tint().orElseThrow())
                      : CellColorReport.theme(theme.theme());
              case ColorSpec.Indexed indexed ->
                  indexed.tint().isPresent()
                      ? CellColorReport.indexed(indexed.indexed(), indexed.tint().orElseThrow())
                      : CellColorReport.indexed(indexed.indexed());
            });
  }
}
