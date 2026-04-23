package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.List;
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
        new ExcelCustomXmlMappingSnapshot(
            1L,
            "CORSO_mapping",
            "CORSO",
            "Schema1",
            false,
            true,
            false,
            true,
            true,
            null,
            null,
            null,
            "<xsd:schema/>",
            null,
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
        () -> new WorkbookCommand.SetArrayFormula(" ", "A1:A2", arrayFormula));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetArrayFormula("Ops", " ", arrayFormula));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearArrayFormula(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearArrayFormula("Ops", " "));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetSignatureLine(" ", signatureDefinition(anchor)));

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
            new ExcelCustomXmlMappingSnapshot(
                0L,
                "CORSO_mapping",
                "CORSO",
                "Schema1",
                false,
                true,
                false,
                true,
                true,
                null,
                null,
                null,
                null,
                null,
                List.of(linkedCell),
                List.of(linkedTable)));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Series(
                null, stringCategories, numericValues, null, null, (short) 1, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Series(
                null, stringCategories, numericValues, null, null, (short) 73, null));
    assertInstanceOf(
        ExcelChartDefinition.Title.None.class,
        new ExcelChartDefinition.Series(null, stringCategories, numericValues).title());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartDefinition.Series(
                null, stringCategories, numericValues, null, null, null, -1L));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.Text("Sales"),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(null, List.of("10.0")),
                null,
                null,
                (short) 1,
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelChartSnapshot.Series(
                null,
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(null, List.of("10.0"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.Text("Sales"),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(null, List.of("10.0")),
                null,
                null,
                null,
                -1L));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelChartSnapshot.Series(
                new ExcelChartSnapshot.Title.Text("Sales"),
                new ExcelChartSnapshot.DataSource.StringLiteral(List.of("Jan")),
                new ExcelChartSnapshot.DataSource.NumericLiteral(null, List.of("10.0")),
                null,
                null,
                (short) 73,
                null));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature", anchor, false, null, null, null, null, null, null, null, null));
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineDefinition(
                "OpsSignature", anchor, false, " ", "Ada", null, null, null, null, null, null));
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
                null,
                null));
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
                null,
                null));
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
                null,
                null));
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
                null,
                null));
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
                null,
                null));
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
                ExcelPictureFormat.PNG,
                null));
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
                null,
                new ExcelBinaryData(new byte[] {1})));
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
        null,
        null);
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
        null,
        null);
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
        null,
        null);
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
        null,
        null);

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelSignatureLineSnapshot(
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
            new ExcelSignatureLineSnapshot(
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
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
    new ExcelDrawingObjectSnapshot.SignatureLine(
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
    new ExcelDrawingObjectSnapshot.SignatureLine(
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
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDrawingObjectSnapshot.SignatureLine(
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
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(new byte[] {1}));
  }
}
