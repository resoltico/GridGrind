package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelSignatureLineSnapshotTestSupport.drawingSignatureLine;
import static dev.erst.gridgrind.excel.ExcelSignatureLineSnapshotTestSupport.signatureLineSnapshot;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlExportSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlLinkedCellSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlLinkedTableSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingSettings;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlMappingSnapshot;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlSchemaSnapshot;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineDefinition;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Residual constructor and validation coverage for rebuilt XLSX value objects and commands. */
class ExcelResidualValidationCoverageTest {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void residualConstructorsRejectInvalidInputs() {
    ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);
    ExcelArrayFormulaDefinition arrayFormula = new ExcelArrayFormulaDefinition("SUM(A1:A2)");
    ExcelCustomXmlMappingLocator locator = new ExcelCustomXmlMappingLocator(1L, "CORSO_mapping");
    ExcelCustomXmlLinkedCellSnapshot linkedCell =
        new ExcelCustomXmlLinkedCellSnapshot("Ops", "A1", "/root/value", "string");
    ExcelCustomXmlLinkedTableSnapshot linkedTable =
        new ExcelCustomXmlLinkedTableSnapshot("Ops", "Table1", "Table 1", "A1:B2", "/root");
    ExcelCustomXmlMappingSnapshot mapping =
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
    ExcelChartDefinition.DataSource.StringLiteral stringCategories =
        new ExcelChartDefinition.DataSource.StringLiteral(List.of("Jan"));
    ExcelChartDefinition.DataSource.NumericLiteral numericValues =
        new ExcelChartDefinition.DataSource.NumericLiteral(List.of(10d));

    assertThrows(IllegalArgumentException.class, () -> new ExcelArrayFormulaDefinition(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelArrayFormulaSnapshot(" ", "A1:A2", "A1", "SUM(A1:A2)", false));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetArrayFormula(" ", "A1:A2", arrayFormula));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.SetArrayFormula("Ops", " ", arrayFormula));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCellCommand.ClearArrayFormula(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCellCommand.ClearArrayFormula("Ops", " "));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookDrawingCommand.SetSignatureLine(" ", signatureDefinition(anchor)));

    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCustomXmlImportDefinition(locator, " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCustomXmlExportSnapshot(mapping, " ", false, "<CORSO/>"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCustomXmlLinkedCellSnapshot(" ", "A1", "/root/value", "string"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCustomXmlLinkedTableSnapshot(" ", "Table1", "Table 1", "A1:B2", "/root"));
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
            new ExcelChartDefinition.Series(
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
            new ExcelChartDefinition.Series(
                null,
                stringCategories,
                numericValues,
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 73),
                Optional.empty()));
    assertInstanceOf(
        ExcelChartDefinition.Title.None.class,
        new ExcelChartDefinition.Series(null, stringCategories, numericValues).title());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Series(
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
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.Text("Sales"),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(Optional.empty(), List.of("10.0")),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 1),
                Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelChartSnapshot.Series(
                null,
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(
                    Optional.empty(), List.of("10.0"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.Text("Sales"),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(Optional.empty(), List.of("10.0")),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1L)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.Text("Sales"),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(Optional.empty(), List.of("10.0")),
                Optional.empty(),
                Optional.empty(),
                Optional.of((short) 73),
                Optional.empty()));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                null,
                null,
                null,
                "one\ntwo\nthree\nfour",
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                " ",
                "Ada",
                null,
                null,
                null,
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                " ",
                null,
                null,
                null,
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                "Ada",
                " ",
                null,
                null,
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                "Ada",
                null,
                " ",
                null,
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                "Ada",
                null,
                null,
                " ",
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                "Ada",
                null,
                null,
                null,
                " ",
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                "Ada",
                null,
                null,
                null,
                null,
                Optional.of(ExcelPictureFormat.PNG),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature",
                anchor,
                false,
                "instructions",
                "Ada",
                null,
                null,
                null,
                null,
                Optional.empty(),
                Optional.of(new ExcelBinaryData(new byte[] {1}))));
    new ExcelSignatureLineDefinition(
        "OpsSignature",
        anchor,
        false,
        "instructions",
        null,
        null,
        null,
        "caption",
        null,
        Optional.empty(),
        Optional.empty());
    new ExcelSignatureLineDefinition(
        "OpsSignature",
        anchor,
        false,
        "instructions",
        null,
        null,
        "ada@example.com",
        null,
        null,
        Optional.empty(),
        Optional.empty());
    new ExcelSignatureLineDefinition(
        "OpsSignature",
        anchor,
        false,
        "instructions",
        null,
        "Finance",
        null,
        null,
        null,
        Optional.empty(),
        Optional.empty());
    new ExcelSignatureLineDefinition(
        "OpsSignature",
        anchor,
        false,
        "instructions",
        null,
        null,
        "ada@example.com",
        "one\ntwo\nthree",
        null,
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

  private static ExcelSignatureLineDefinition signatureDefinition(
      ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelSignatureLineDefinition(
        "OpsSignature",
        anchor,
        false,
        "Review before signing.",
        "Ada Lovelace",
        "Finance",
        "ada@example.com",
        null,
        "invalid",
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(new ExcelBinaryData(new byte[] {1})));
  }

  private static ExcelCustomXmlMappingSnapshot customXmlMappingSnapshot(
      long mapId,
      String name,
      String rootElement,
      String schemaId,
      ExcelCustomXmlMappingSettings settings,
      ExcelCustomXmlSchemaSnapshot schema,
      Optional<dev.erst.gridgrind.excel.customxml.ExcelCustomXmlDataBindingSnapshot> dataBinding,
      List<ExcelCustomXmlLinkedCellSnapshot> linkedCells,
      List<ExcelCustomXmlLinkedTableSnapshot> linkedTables) {
    return new ExcelCustomXmlMappingSnapshot(
        mapId,
        name,
        rootElement,
        schemaId,
        settings,
        schema,
        dataBinding,
        linkedCells,
        linkedTables);
  }

  private static ExcelCustomXmlMappingSettings settings(
      boolean showImportExportValidationErrors,
      boolean autoFit,
      boolean append,
      boolean preserveSortAfLayout,
      boolean preserveFormat) {
    return new ExcelCustomXmlMappingSettings(
        showImportExportValidationErrors, autoFit, append, preserveSortAfLayout, preserveFormat);
  }

  private static ExcelCustomXmlSchemaSnapshot schema(
      Optional<String> namespace,
      Optional<String> language,
      Optional<String> reference,
      Optional<String> xml) {
    return new ExcelCustomXmlSchemaSnapshot(namespace, language, reference, xml);
  }
}
