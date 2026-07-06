package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptorWithNotes;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;

import dev.erst.gridgrind.contract.dto.CellReport;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.CellValueReport;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.NamedRangeReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.SheetProtectionReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WindowReport;
import dev.erst.gridgrind.contract.query.CellReadFacet;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Arrays;
import java.util.List;

/** Owns one focused subset of nested-type group descriptors for the protocol catalog. */
final class GridGrindProtocolCatalogSourceAndReportNestedTypeGroups {
  private GridGrindProtocolCatalogSourceAndReportNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> SOURCE_AND_REPORT_GROUPS =
      List.of(
          nestedTypeGroup(
              "cellScalarValueTypes",
              CellScalarValue.class,
              List.of(
                  descriptor(
                      CellScalarValue.Blank.class,
                      "BLANK",
                      "Require the effective cell value to be blank."),
                  descriptor(
                      CellScalarValue.Text.class,
                      "TEXT",
                      "Require the effective cell value to be one exact string."),
                  descriptor(
                      CellScalarValue.NumberValue.class,
                      "NUMBER",
                      "Require the effective cell value to be one exact finite number."),
                  descriptor(
                      CellScalarValue.BooleanValue.class,
                      "BOOLEAN",
                      "Require the effective cell value to be true or false."),
                  descriptor(
                      CellScalarValue.ErrorValue.class,
                      "ERROR",
                      "Require the effective cell value to be one exact GridGrind-owned error"
                          + " literal, including canonical Excel tokens and evaluation-only"
                          + " states such as #CIRCULAR_REF!."))),
          nestedTypeGroup(
              "cellReportTypes",
              CellReport.class,
              List.of(
                  descriptor(
                      CellReport.BlankReport.class,
                      "BLANK",
                      "Factual blank cell report. Only address and type are always present;"
                          + " FORMAT, STYLE, HYPERLINK, and COMMENT project the remaining fields.",
                      List.of(),
                      projectedField("displayValue", CellReadFacet.FORMAT),
                      projectedField("style", CellReadFacet.STYLE),
                      projectedField("hyperlink", CellReadFacet.HYPERLINK),
                      projectedField("comment", CellReadFacet.COMMENT)),
                  descriptor(
                      CellReport.TextReport.class,
                      "TEXT",
                      "Factual text cell report. VALUE projects textValue,"
                          + " RICH_TEXT_RUNS projects runs, and the remaining fields are"
                          + " facet-gated.",
                      List.of(),
                      projectedField("displayValue", CellReadFacet.FORMAT),
                      projectedField("style", CellReadFacet.STYLE),
                      projectedField("hyperlink", CellReadFacet.HYPERLINK),
                      projectedField("comment", CellReadFacet.COMMENT),
                      projectedField("textValue", CellReadFacet.VALUE),
                      projectedField("runs", CellReadFacet.RICH_TEXT_RUNS)),
                  descriptor(
                      CellReport.NumberReport.class,
                      "NUMBER",
                      "Factual numeric cell report. VALUE projects numberValue; TEMPORAL adds"
                          + " date, time, or date-time semantics when a numeric Excel value is"
                          + " format-derived rather than a plain number.",
                      List.of(),
                      projectedField("displayValue", CellReadFacet.FORMAT),
                      projectedField("style", CellReadFacet.STYLE),
                      projectedField("hyperlink", CellReadFacet.HYPERLINK),
                      projectedField("comment", CellReadFacet.COMMENT),
                      projectedField("numberValue", CellReadFacet.VALUE),
                      projectedField("temporal", CellReadFacet.TEMPORAL)),
                  descriptor(
                      CellReport.BooleanReport.class,
                      "BOOLEAN",
                      "Factual boolean cell report. VALUE projects booleanValue and the"
                          + " remaining fields are facet-gated.",
                      List.of(),
                      projectedField("displayValue", CellReadFacet.FORMAT),
                      projectedField("style", CellReadFacet.STYLE),
                      projectedField("hyperlink", CellReadFacet.HYPERLINK),
                      projectedField("comment", CellReadFacet.COMMENT),
                      projectedField("booleanValue", CellReadFacet.VALUE)),
                  descriptor(
                      CellReport.ErrorReport.class,
                      "ERROR",
                      "Factual error cell report. VALUE projects errorValue and the remaining"
                          + " fields are facet-gated.",
                      List.of(),
                      projectedField("displayValue", CellReadFacet.FORMAT),
                      projectedField("style", CellReadFacet.STYLE),
                      projectedField("hyperlink", CellReadFacet.HYPERLINK),
                      projectedField("comment", CellReadFacet.COMMENT),
                      projectedField("errorValue", CellReadFacet.VALUE)),
                  descriptor(
                      CellReport.FormulaReport.class,
                      "FORMULA",
                      "Factual formula cell report with separate projected evaluation."
                          + " FORMULA projects formula text, VALUE projects evaluation, and the"
                          + " remaining fields are facet-gated.",
                      List.of(),
                      projectedField("displayValue", CellReadFacet.FORMAT),
                      projectedField("style", CellReadFacet.STYLE),
                      projectedField("hyperlink", CellReadFacet.HYPERLINK),
                      projectedField("comment", CellReadFacet.COMMENT),
                      projectedField("formula", CellReadFacet.FORMULA),
                      projectedField("evaluation", CellReadFacet.VALUE)))),
          nestedTypeGroup(
              "cellValueReportTypes",
              CellValueReport.class,
              List.of(
                  descriptor(
                      CellValueReport.BlankValue.class,
                      "BLANK",
                      "Effective formula evaluation is blank."),
                  descriptor(
                      CellValueReport.TextValue.class,
                      "TEXT",
                      "Effective formula evaluation is text; RICH_TEXT_RUNS projects optional"
                          + " runs.",
                      List.of(),
                      projectedField("runs", CellReadFacet.RICH_TEXT_RUNS)),
                  descriptor(
                      CellValueReport.NumberValue.class,
                      "NUMBER",
                      "Effective formula evaluation is numeric; TEMPORAL projects derived"
                          + " date, time, or date-time semantics when formatting marks the"
                          + " numeric value as date-like.",
                      List.of(),
                      projectedField("temporal", CellReadFacet.TEMPORAL)),
                  descriptor(
                      CellValueReport.BooleanValue.class,
                      "BOOLEAN",
                      "Effective formula evaluation is boolean."),
                  descriptor(
                      CellValueReport.ErrorValue.class,
                      "ERROR",
                      "Effective formula evaluation is one GridGrind-owned error literal,"
                          + " including canonical Excel tokens plus evaluator-specific tokens"
                          + " such as #CIRCULAR_REF!."))),
          nestedTypeGroup(
              "windowReportTypes",
              WindowReport.class,
              List.of(
                  descriptor(
                      WindowReport.Sparse.class,
                      "SPARSE",
                      "Sparse window omitting blank cells and publishing populatedCells only."),
                  descriptor(
                      WindowReport.Dense.class,
                      "DENSE",
                      "Dense window preserving explicit row structure including blank cells."))),
          nestedTypeGroup(
              "textSourceTypes",
              TextSourceInput.class,
              List.of(
                  descriptor(
                      TextSourceInput.Inline.class,
                      "INLINE",
                      "Embed UTF-8 text directly in the request JSON."),
                  descriptorWithNotes(
                      TextSourceInput.Utf8File.class,
                      "UTF8_FILE",
                      "Load UTF-8 text from one file path.",
                      GridGrindProtocolCatalogNotes.requestOwnedPathRuleRef()),
                  descriptor(
                      TextSourceInput.StandardInput.class,
                      "STANDARD_INPUT",
                      "Load UTF-8 text from the execution transport's bound standard input"
                          + " bytes."))),
          nestedTypeGroup(
              "binarySourceTypes",
              BinarySourceInput.class,
              List.of(
                  descriptor(
                      BinarySourceInput.InlineBase64.class,
                      "INLINE_BASE64",
                      "Embed base64-encoded binary content directly in the request JSON."),
                  descriptorWithNotes(
                      BinarySourceInput.File.class,
                      "FILE",
                      "Load binary content from one file path.",
                      GridGrindProtocolCatalogNotes.requestOwnedPathRuleRef()),
                  descriptor(
                      BinarySourceInput.StandardInput.class,
                      "STANDARD_INPUT",
                      "Load binary content from the execution transport's bound standard input"
                          + " bytes."))),
          nestedTypeGroup(
              "namedRangeReportTypes",
              NamedRangeReport.class,
              List.of(
                  descriptor(
                      NamedRangeReport.RangeReport.class,
                      "RANGE",
                      "Exact named-range report that resolves to one typed workbook target."),
                  descriptor(
                      NamedRangeReport.FormulaReport.class,
                      "FORMULA",
                      "Exact named-range report that remains formula-backed."))),
          nestedTypeGroup(
              "sheetProtectionReportTypes",
              SheetProtectionReport.class,
              List.of(
                  descriptor(
                      SheetProtectionReport.Unprotected.class,
                      "UNPROTECTED",
                      "Expect the sheet to have no protection."),
                  descriptor(
                      SheetProtectionReport.Protected.class,
                      "PROTECTED",
                      "Expect the sheet to be protected with explicit lock settings."))),
          nestedTypeGroup(
              "tableStyleReportTypes",
              TableStyleReport.class,
              List.of(
                  descriptor(
                      TableStyleReport.None.class,
                      "NONE",
                      "Expect the table to carry no persisted style."),
                  descriptor(
                      TableStyleReport.Named.class,
                      "NAMED",
                      "Expect the table to carry one named style plus stripe/emphasis flags."))),
          nestedTypeGroup(
              "pivotTableReportTypes",
              PivotTableReport.class,
              List.of(
                  descriptor(
                      PivotTableReport.Supported.class,
                      "SUPPORTED",
                      "Exact supported pivot-table report."),
                  descriptor(
                      PivotTableReport.Unsupported.class,
                      "UNSUPPORTED",
                      "Exact unsupported pivot-table report preserved from the workbook."))),
          nestedTypeGroup(
              "pivotTableReportSourceTypes",
              PivotTableReport.Source.class,
              List.of(
                  descriptor(
                      PivotTableReport.Source.Range.class,
                      "RANGE",
                      "Pivot source resolved from one sheet range."),
                  descriptor(
                      PivotTableReport.Source.NamedRange.class,
                      "NAMED_RANGE",
                      "Pivot source resolved from one named range."),
                  descriptor(
                      PivotTableReport.Source.Table.class,
                      "TABLE",
                      "Pivot source resolved from one workbook table."))),
          nestedTypeGroup(
              "chartPlotReportTypes",
              ChartReport.Plot.class,
              List.of(
                  descriptor(ChartReport.Area.class, "AREA", "Exact area-chart plot report."),
                  descriptor(
                      ChartReport.Area3D.class,
                      "AREA_3D",
                      "Exact 3D area-chart plot report.",
                      "gapDepth"),
                  descriptor(
                      ChartReport.Bar.class,
                      "BAR",
                      "Exact bar-chart plot report.",
                      "gapWidth",
                      "overlap"),
                  descriptor(
                      ChartReport.Bar3D.class,
                      "BAR_3D",
                      "Exact 3D bar-chart plot report.",
                      "gapDepth",
                      "gapWidth",
                      "shape"),
                  descriptor(
                      ChartReport.Doughnut.class,
                      "DOUGHNUT",
                      "Exact doughnut-chart plot report.",
                      "firstSliceAngle",
                      "holeSize"),
                  descriptor(ChartReport.Line.class, "LINE", "Exact line-chart plot report."),
                  descriptor(
                      ChartReport.Line3D.class,
                      "LINE_3D",
                      "Exact 3D line-chart plot report.",
                      "gapDepth"),
                  descriptor(
                      ChartReport.Pie.class,
                      "PIE",
                      "Exact pie-chart plot report.",
                      "firstSliceAngle"),
                  descriptor(ChartReport.Pie3D.class, "PIE_3D", "Exact 3D pie-chart plot report."),
                  descriptor(ChartReport.Radar.class, "RADAR", "Exact radar-chart plot report."),
                  descriptor(
                      ChartReport.Scatter.class, "SCATTER", "Exact scatter-chart plot report."),
                  descriptor(
                      ChartReport.Surface.class, "SURFACE", "Exact surface-chart plot report."),
                  descriptor(
                      ChartReport.Surface3D.class,
                      "SURFACE_3D",
                      "Exact 3D surface-chart plot report."),
                  descriptor(
                      ChartReport.Unsupported.class,
                      "UNSUPPORTED",
                      "Exact unsupported chart plot report preserved from the workbook."))),
          nestedTypeGroup(
              "chartTitleReportTypes",
              ChartReport.Title.class,
              List.of(
                  descriptor(ChartReport.Title.None.class, "NONE", "No title is present."),
                  descriptor(ChartReport.Title.Text.class, "TEXT", "Static title text."),
                  descriptor(
                      ChartReport.Title.Formula.class,
                      "FORMULA",
                      "Formula-backed title with cached text."))),
          nestedTypeGroup(
              "chartLegendReportTypes",
              ChartReport.Legend.class,
              List.of(
                  descriptor(ChartReport.Legend.Hidden.class, "HIDDEN", "No legend is present."),
                  descriptor(
                      ChartReport.Legend.Visible.class,
                      "VISIBLE",
                      "Visible legend at one persisted position."))),
          nestedTypeGroup(
              "chartDataSourceReportTypes",
              ChartReport.DataSource.class,
              List.of(
                  descriptor(
                      ChartReport.DataSource.StringReference.class,
                      "STRING_REFERENCE",
                      "Formula-backed string chart source plus cached values."),
                  descriptor(
                      ChartReport.DataSource.NumericReference.class,
                      "NUMERIC_REFERENCE",
                      "Formula-backed numeric chart source plus cached values."),
                  descriptor(
                      ChartReport.DataSource.StringLiteral.class,
                      "STRING_LITERAL",
                      "Literal string chart source stored directly in the chart part."),
                  descriptor(
                      ChartReport.DataSource.NumericLiteral.class,
                      "NUMERIC_LITERAL",
                      "Literal numeric chart source stored directly in the chart part."))),
          nestedTypeGroup(
              "drawingAnchorReportTypes",
              DrawingAnchorReport.class,
              List.of(
                  descriptor(
                      DrawingAnchorReport.TwoCell.class,
                      "TWO_CELL",
                      "Drawing anchor spanning one start and end marker."),
                  descriptor(
                      DrawingAnchorReport.OneCell.class,
                      "ONE_CELL",
                      "Drawing anchor with one start marker plus explicit size."),
                  descriptor(
                      DrawingAnchorReport.Absolute.class,
                      "ABSOLUTE",
                      "Drawing anchor with absolute EMU coordinates and size."))));

  private static CatalogProjectedField projectedField(String name, CellReadFacet firstFacet) {
    return projectedField(name, firstFacet, new CellReadFacet[0]);
  }

  private static CatalogProjectedField projectedField(
      String name, CellReadFacet firstFacet, CellReadFacet... remainingFacets) {
    return new CatalogProjectedField(name, facetNames(firstFacet, remainingFacets));
  }

  private static List<String> facetNames(
      CellReadFacet firstFacet, CellReadFacet... remainingFacets) {
    return java.util.stream.Stream.concat(
            java.util.stream.Stream.of(firstFacet), Arrays.stream(remainingFacets))
        .map(CellReadFacet::name)
        .toList();
  }
}
