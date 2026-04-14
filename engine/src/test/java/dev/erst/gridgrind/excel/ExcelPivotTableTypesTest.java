package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.junit.jupiter.api.Test;

/** Tests for pivot-table value objects, validation helpers, and enum mappings. */
class ExcelPivotTableTypesTest {
  @Test
  void validatesNamingSelectionsDefinitionsAndSnapshots() {
    assertEquals("Sales Pivot 2026", ExcelPivotTableNaming.validateName("Sales Pivot 2026"));
    assertThrows(NullPointerException.class, () -> ExcelPivotTableNaming.validateName(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelPivotTableNaming.validateName("   "));

    List<String> mutableNames = new ArrayList<>(List.of("Sales Pivot 2026", "Ops Pivot"));
    ExcelPivotTableSelection.ByNames selection = new ExcelPivotTableSelection.ByNames(mutableNames);
    mutableNames.clear();
    assertEquals(List.of("Sales Pivot 2026", "Ops Pivot"), selection.names());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSelection.ByNames(List.of("Sales Pivot 2026", "sales pivot 2026")));

    ExcelPivotTableDefinition definition =
        new ExcelPivotTableDefinition(
            "Sales Pivot 2026",
            "Report",
            new ExcelPivotTableDefinition.Source.Table("SalesTable"),
            new ExcelPivotTableDefinition.Anchor("C5"),
            List.of("Region"),
            List.of("Stage"),
            List.of("Owner"),
            List.of(
                new ExcelPivotTableDefinition.DataField(
                    "Amount", ExcelPivotDataConsolidateFunction.SUM, null, "#,##0.00")));

    assertEquals("Sales Pivot 2026", definition.name());
    assertEquals("Amount", definition.dataFields().getFirst().displayName());
    assertEquals("#,##0.00", definition.dataFields().getFirst().valueFormat());
    assertEquals(
        "Amount",
        new ExcelPivotTableDefinition.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, " ", null)
            .displayName());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPivotTableDefinition.Source.NamedRange("Sales Pivot 2026"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPivotTableDefinition.Source.Table("Sales Table With Spaces"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Sales Pivot 2026",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A1"),
                List.of("Region"),
                List.of("Stage"),
                List.of("Owner"),
                List.of(
                    new ExcelPivotTableDefinition.DataField(
                        "Region", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableDefinition.DataField(
                "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", " "));

    ExcelPivotTableSnapshot.Supported supported =
        new ExcelPivotTableSnapshot.Supported(
            "Sales Pivot 2026",
            "Report",
            new ExcelPivotTableSnapshot.Anchor("C5", "C5:F9"),
            new ExcelPivotTableSnapshot.Source.Range("Data", "A1:D5"),
            List.of(new ExcelPivotTableSnapshot.Field(0, "Region")),
            List.of(new ExcelPivotTableSnapshot.Field(1, "Stage")),
            List.of(new ExcelPivotTableSnapshot.Field(2, "Owner")),
            List.of(
                new ExcelPivotTableSnapshot.DataField(
                    3,
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    "#,##0.00")),
            true);
    ExcelPivotTableSnapshot.Unsupported unsupported =
        new ExcelPivotTableSnapshot.Unsupported(
            "Broken Pivot",
            "Report",
            new ExcelPivotTableSnapshot.Anchor("A3", "A3:C6"),
            "Pivot cache source no longer resolves cleanly.");

    assertTrue(supported.valuesAxisOnColumns());
    assertEquals("Amount", supported.dataFields().getFirst().sourceColumnName());
    assertEquals("Pivot cache source no longer resolves cleanly.", unsupported.detail());
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPivotTableSnapshot.Field(-1, "Region"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                0, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", " "));
  }

  @Test
  void pivotDataConsolidateFunctionsRoundTripSupportedPoiValues() {
    assertEquals(
        DataConsolidateFunction.SUM, ExcelPivotDataConsolidateFunction.SUM.toPoiFunction());
    assertEquals(
        ExcelPivotDataConsolidateFunction.MAX,
        ExcelPivotDataConsolidateFunction.fromPoiFunction(DataConsolidateFunction.MAX));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelPivotDataConsolidateFunction.fromPoiFunction(null));
  }

  @Test
  void pivotTypesRejectAdditionalInvalidInputs() {
    assertThrows(NullPointerException.class, () -> new ExcelPivotTableDefinition.Anchor(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelPivotTableDefinition.Anchor(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPivotTableDefinition.Source.Range("Data", " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableDefinition.DataField(
                null, ExcelPivotDataConsolidateFunction.SUM, "Total", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableDefinition.DataField(
                " ", ExcelPivotDataConsolidateFunction.SUM, "Total", null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelPivotTableDefinition.DataField("Amount", null, "Total", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Missing Data",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A3"),
                null,
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Empty Data",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A3"),
                null,
                null,
                null,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Null Data Field",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A3"),
                null,
                null,
                null,
                new ArrayList<>(
                    java.util.Arrays.asList((ExcelPivotTableDefinition.DataField) null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Axis Clash",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A3"),
                List.of("Owner"),
                List.of("Region"),
                List.of("region"),
                List.of(
                    new ExcelPivotTableDefinition.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Duplicate Axis Names",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A3"),
                List.of(),
                List.of("Region", "region"),
                List.of(),
                List.of(
                    new ExcelPivotTableDefinition.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Null Axis Name",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A3"),
                new ArrayList<>(java.util.Arrays.asList((String) null)),
                List.of(),
                List.of(),
                List.of(
                    new ExcelPivotTableDefinition.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableDefinition(
                "Blank Axis",
                "Report",
                new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                new ExcelPivotTableDefinition.Anchor("A3"),
                List.of(" "),
                List.of(),
                List.of(),
                List.of(
                    new ExcelPivotTableDefinition.DataField(
                        "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total", null))));

    assertThrows(NullPointerException.class, () -> new ExcelPivotTableSelection.ByNames(null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelPivotTableSelection.ByNames(List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSelection.ByNames(
                new ArrayList<>(java.util.Arrays.asList("Sales Pivot 2026", null))));

    ExcelPivotTableSnapshot.Anchor anchor = new ExcelPivotTableSnapshot.Anchor("C5", "C5:F9");
    ExcelPivotTableSnapshot.Source.Range source =
        new ExcelPivotTableSnapshot.Source.Range("Data", "A1:D5");
    ExcelPivotTableSnapshot.Field rowField = new ExcelPivotTableSnapshot.Field(0, "Region");
    ExcelPivotTableSnapshot.DataField dataField =
        new ExcelPivotTableSnapshot.DataField(
            3, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null);

    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.Supported(
                "Sales Pivot 2026",
                "Report",
                anchor,
                source,
                null,
                List.of(),
                List.of(),
                List.of(dataField),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.Supported(
                "Sales Pivot 2026",
                "Report",
                anchor,
                source,
                List.of(rowField),
                List.of(),
                List.of(),
                new ArrayList<>(java.util.Arrays.asList((ExcelPivotTableSnapshot.DataField) null)),
                false));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, "Amount", ExcelPivotDataConsolidateFunction.SUM, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, "Amount", ExcelPivotDataConsolidateFunction.SUM, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelPivotTableSnapshot.Unsupported("Broken Pivot", "Report", anchor, " "));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, null, ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                3, " ", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelPivotTableSnapshot.DataField(3, "Amount", null, "Total Amount", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelPivotTableSnapshot.DataField(
                -1, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", null));
  }
}
