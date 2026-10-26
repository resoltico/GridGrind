package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelSignatureLineSnapshotTestSupport.drawingSignatureLine;
import static dev.erst.gridgrind.excel.ExcelSignatureLineSnapshotTestSupport.signatureLineSnapshot;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Residual constructor and validation coverage for rebuilt XLSX value objects and commands. */
class ExcelResidualValidationCoverageTest {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void residualConstructorsRejectInvalidInputs() {
    DrawingAnchorSpec.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);
    ResolvedArrayFormula arrayFormula = new ResolvedArrayFormula("SUM(A1:A2)");
    CustomXmlMappingLocator locator =
        new CustomXmlMappingLocator(
            java.util.Optional.of(1L), java.util.Optional.of("CORSO_mapping"));
    CustomXmlLinkedCellReport linkedCell =
        new CustomXmlLinkedCellReport(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
            new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
            "/root/value",
            "string");
    CustomXmlLinkedTableReport linkedTable =
        new CustomXmlLinkedTableReport(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
            "Table1",
            "Table 1",
            new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"),
            "/root");
    CustomXmlMappingReport mapping =
        customXmlMappingSnapshot(
            1L,
            "CORSO_mapping",
            "CORSO",
            "Schema1",
            settings(false, true, false, true, true),
            schema(
                Optional.empty(), Optional.empty(), Optional.empty(), Optional.of("<xsd:schema/>")),
            Optional.empty(),
            List.of(linkedCell),
            List.of(linkedTable));
    ChartDataSourceSpec.StringLiteral stringCategories =
        new ChartDataSourceSpec.StringLiteral(List.of("Jan"));
    ChartDataSourceSpec.NumericLiteral numericValues =
        new ChartDataSourceSpec.NumericLiteral(List.of(10d));

    assertThrows(IllegalArgumentException.class, () -> new ResolvedArrayFormula(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ArrayFormulaReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(" "),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:A2"),
                new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                "SUM(A1:A2)",
                false));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedStep(
                new RangeSelector.ByRange(new SheetName(" "), new RangeRef("A1:A2")),
                new ResolvedMutationAction.SetArrayFormula(arrayFormula)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedStep(
                new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef(" ")),
                new ResolvedMutationAction.SetArrayFormula(arrayFormula)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedStep(
                new CellSelector.ByAddress(new SheetName(" "), new CellAddress("A1")),
                new CellMutationAction.ClearArrayFormula()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedStep(
                new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress(" ")),
                new CellMutationAction.ClearArrayFormula()));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedStep(
                new SheetSelector.ByName(new SheetName(" ")),
                new ResolvedMutationAction.SetSignatureLine(signatureDefinition(anchor))));

    assertThrows(IllegalArgumentException.class, () -> new ResolvedCustomXmlImport(locator, " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CustomXmlExportReport(mapping, " ", false, "<CORSO/>"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CustomXmlLinkedCellReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(" "),
                new dev.erst.gridgrind.core.vocabulary.CellAddress("A1"),
                "/root/value",
                "string"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CustomXmlLinkedTableReport(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(" "),
                "Table1",
                "Table 1",
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"),
                "/root"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            customXmlMappingSnapshot(
                0L,
                "CORSO_mapping",
                "CORSO",
                "Schema1",
                settings(false, true, false, true, true),
                schema(null, null, null, null),
                null,
                List.of(linkedCell),
                List.of(linkedTable)));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedChart.Series(
                null,
                stringCategories,
                numericValues,
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedChart.Series(
                null,
                stringCategories,
                numericValues,
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 73),
                Optional.empty()));
    assertInstanceOf(
        ChartTitleSpec.None.class,
        new ResolvedChart.Series(null, stringCategories, numericValues).title());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedChart.Series(
                null,
                stringCategories,
                numericValues,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1L)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Sales"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10.0")),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 1),
                Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ChartReport.Series(
                null,
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10.0")),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Sales"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10.0")),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1L)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ChartReport.Series(
                new ChartReport.Title.Text("Sales"),
                new ChartReport.DataSource.StringLiteral(List.of("Jan")),
                new ChartReport.DataSource.NumericLiteral(Optional.empty(), List.of("10.0")),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 73),
                Optional.empty()));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable("one\ntwo\nthree\nfour"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable(" "),
                Optional.ofNullable("Ada"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.ofNullable(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.empty(),
                Optional.ofNullable(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable(" "),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(ExcelPictureFormat.PNG),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedSignatureLine(
                "OpsSignature",
                anchor,
                false,
                Optional.ofNullable("instructions"),
                Optional.ofNullable("Ada"),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(new BinaryContent(new byte[] {1}))));
    new ResolvedSignatureLine(
        "OpsSignature",
        anchor,
        false,
        Optional.ofNullable("instructions"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.ofNullable("caption"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
    new ResolvedSignatureLine(
        "OpsSignature",
        anchor,
        false,
        Optional.ofNullable("instructions"),
        Optional.empty(),
        Optional.empty(),
        Optional.ofNullable("ada@example.com"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
    new ResolvedSignatureLine(
        "OpsSignature",
        anchor,
        false,
        Optional.ofNullable("instructions"),
        Optional.empty(),
        Optional.ofNullable("Finance"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
    new ResolvedSignatureLine(
        "OpsSignature",
        anchor,
        false,
        Optional.ofNullable("instructions"),
        Optional.empty(),
        Optional.empty(),
        Optional.ofNullable("ada@example.com"),
        Optional.ofNullable("one\ntwo\nthree"),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                "image/png",
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
                "OpsSignature",
                anchor,
                null,
                false,
                " ",
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                " ",
                null,
                null,
                null,
                null,
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                4L,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                " ",
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                ExcelPictureFormat.PNG,
                "image/png",
                4L,
                "hash",
                -1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            signatureLineSnapshot(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                ExcelPictureFormat.PNG,
                "image/png",
                4L,
                "hash",
                null,
                -1));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                "image/png",
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                " ",
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                " ",
                null,
                null,
                null,
                null,
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                " ",
                null,
                null,
                null,
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                " ",
                null,
                null,
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                4L,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                " ",
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                ExcelPictureFormat.PNG,
                "image/png",
                4L,
                "hash",
                -1,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                ExcelPictureFormat.PNG,
                "image/png",
                4L,
                "hash",
                null,
                -1));
    drawingSignatureLine(
        "OpsSignature",
        anchor,
        null,
        false,
        null,
        null,
        null,
        null,
        ExcelPictureFormat.PNG,
        "image/png",
        4L,
        "hash",
        null,
        0);
    drawingSignatureLine(
        "OpsSignature",
        anchor,
        null,
        false,
        null,
        null,
        null,
        null,
        ExcelPictureFormat.PNG,
        "image/png",
        4L,
        "hash",
        null,
        null);
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                " ",
                null,
                null,
                -1,
                -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            drawingSignatureLine(
                "OpsSignature",
                anchor,
                null,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                -1));
  }

  private static ResolvedSignatureLine signatureDefinition(DrawingAnchorSpec.TwoCell anchor) {
    return new ResolvedSignatureLine(
        "OpsSignature",
        anchor,
        false,
        Optional.ofNullable("Review before signing."),
        Optional.ofNullable("Ada Lovelace"),
        Optional.ofNullable("Finance"),
        Optional.ofNullable("ada@example.com"),
        Optional.empty(),
        Optional.ofNullable("invalid"),
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(new BinaryContent(new byte[] {1})));
  }

  private static CustomXmlMappingReport customXmlMappingSnapshot(
      long mapId,
      String name,
      String rootElement,
      String schemaId,
      CustomXmlMappingReport.Settings settings,
      CustomXmlMappingReport.Schema schema,
      Optional<dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport> dataBinding,
      List<CustomXmlLinkedCellReport> linkedCells,
      List<CustomXmlLinkedTableReport> linkedTables) {
    return new CustomXmlMappingReport(
        mapId,
        name,
        rootElement,
        schemaId,
        settings,
        schema,
        dataBinding == null ? null : dataBinding.orElse(null),
        linkedCells,
        linkedTables);
  }

  private static CustomXmlMappingReport.Settings settings(
      boolean showImportExportValidationErrors,
      boolean autoFit,
      boolean append,
      boolean preserveSortAfLayout,
      boolean preserveFormat) {
    return new CustomXmlMappingReport.Settings(
        showImportExportValidationErrors, autoFit, append, preserveSortAfLayout, preserveFormat);
  }

  private static CustomXmlMappingReport.Schema schema(
      Optional<String> namespace,
      Optional<String> language,
      Optional<String> reference,
      Optional<String> xml) {
    return new CustomXmlMappingReport.Schema(
        nullable(namespace), nullable(language), nullable(reference), nullable(xml));
  }

  private static <T> T nullable(Optional<T> value) {
    return value == null ? null : value.orElse(null);
  }
}
